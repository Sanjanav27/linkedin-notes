from flask import Flask,render_template,request,url_for,redirect
import numpy as np
import pandas as pd
from pymongo import MongoClient
from dotenv import load_dotenv
import os
from openpyxl import Workbook
import xlwt
from xlwt import Workbook
load_dotenv()



app = Flask(__name__)


  
# Workbook is created
wb = Workbook()
  
# add_sheet is used to create sheet.
sheet2 = wb.add_sheet('Sheet2')
  



MONGO_URI = os.getenv("MONGO_URI")
cluster=MongoClient(MONGO_URI)
db=cluster["linkedin-notes"]
collection=db["detail"]



data = pd.read_excel (r'static/LinkedIn_Notes.xlsx',sheet_name='Sheet1') 

df = pd.DataFrame(data, columns= ['Name','LinkedIn link'])

df.dropna(inplace = True)
Linked_name = df['Name'].tolist() 
ll_link=df['LinkedIn link'].tolist()

dic=df.to_dict()
print(dic)

size=len(Linked_name)

def page(pg , index):
    
    if pg == 'next':
        if size-1 > index:
            return Linked_name[index+1]
        elif size-1 == index:
            return Linked_name[0]
    elif pg == 'pre':
        if index == 0 :
            return Linked_name[size-1]
        else:
            return Linked_name[index-1]

@app.route('/')
@app.route('/crowdengine/')
def crowdengine():
    return render_template('mainpage.html',movie_list = Linked_name,link=ll_link)

@app.route('/pyhackons/')
def pyhackons():
    return render_template('pyhackons.html')

@app.route('/crowdengine/<string:name>')
def movie_name(name):
    
    if name not in Linked_name:
        return render_template('error.html',error="Name Not Found")
    x=Linked_name.index(name)
    link=ll_link[x]
    print(link)
    namemongo=[]
    for all in collection.find({},{ "_id": 0,"name":1}):
        print(all)
        for all in all.values():
            namemongo.append(all)
    print(namemongo)
    a=[]
    if name in namemongo:
        query={"name":name}
        doc=collection.find(query)
        a = []
        for y in doc:
            print(y)
            a.append(y["notes"])
        
    
    return render_template('movies.html' ,name = name,link=link,a=a)

@app.errorhandler(500)
def internal_error(e):
    return render_template('error.html',error='Problem occured')

@app.errorhandler(404)
def page_not_found(e):
    
    return render_template('error.html',error='Page Not found')

@app.route('/crowdengine/pre/')
def pre():
    name = request.args.get('page')
    print(name)
    index = Linked_name.index(name)
    name = page('pre',index)
    x=Linked_name.index(name)
    link=ll_link[x]

    return redirect(url_for('movie_name',name = name,link=link) )
    
   

@app.route('/crowdengine/next/')
def next():
    name = request.args.get('page')
    print(name)
    index = Linked_name.index(name)
    name = page('next',index)
    x=Linked_name.index(name)
    link=ll_link[x]
    print(link)
    return redirect(url_for('movie_name',name = name,link=link))

@app.route('/crowdengine/write/<name>',methods=["POST"])   
def write_db(name):

    
    if request.method == 'POST':    
        x=Linked_name.index(name)
        link=ll_link[x]
        movie = request.form["movie"]
        actor = request.form["name"]
        dur = request.form["Duration"]
        hair = request.form["Hairstyle"]
        role = request.form["Role"]
        hit = request.form["Hit"]
        color = request.form["Dresscolor"]
        print(link)
        comment = request.form["logname"]
        print(comment)
        
        data.write(movie=movie,actor=actor,duration=dur,hairstyle=hair,role=role,dresscolor=color,target=hit)

    return render_template('movies.html',p="ok" ,name=movie,link="link")

@app.route('/download/')
def download():
    return data.get_csv(a = app)

@app.route('/submit/<string:name>',methods=["POST","GET"])
def submit(name):
    name=name
    comment=request.form.get("comment")
    print(comment)
    x=Linked_name.index(name)
    link=ll_link[x]
    namemongo=[]
    for all in collection.find({},{ "_id": 0,"name":1}):
        print(all)
        for all in all.values():
            namemongo.append(all)
    print(namemongo)
    a=[]
    if name in namemongo:
        query={"name":name}
        newvalues = { "$set": { "notes": comment } }

        collection.update_one(query, newvalues)
        sheet2.write(x, 1, comment)  


    else:
        data={
        "name" : name,
        "link" : link,
        "notes":comment
        }


        collection.insert_one(data)
        sheet2.write(x, 0, name) 
        sheet2.write(x, 2, comment)  
        wb.save('static/data.xlsx')
    


    return "added successfully"

if __name__ == "__main__":
    app.run()
    # a=get_xl_data()

    
