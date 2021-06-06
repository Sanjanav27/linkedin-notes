from flask import Flask,render_template,request,url_for,redirect
import numpy as np
import pandas as pd


app = Flask(__name__ )



data = pd.read_excel (r'static/LinkedIn_Notes.xlsx',sheet_name='Sheet1') 

df = pd.DataFrame(data, columns= ['Name','LinkedIn link'])

df.dropna(inplace = True)
Linked_name = df['Name'].tolist() 

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
    return render_template('mainpage.html',movie_list = Linked_name)

@app.route('/pyhackons/')
def pyhackons():
    return render_template('pyhackons.html')

@app.route('/crowdengine/<string:name>')
def movie_name(name):
    
    # if name not in Linked_name:
    #     return render_template('error.html',error="Name Not Found")
    return render_template('movies.html' ,name = name,actor ="vadivelu")

@app.errorhandler(500)
def internal_error(e):
    return render_template('error.html',error='Problem occured')

@app.errorhandler(404)
def page_not_found(e):
    
    return render_template('error.html',error='Page Not found')

@app.route('/crowdengine/pre/')
def pre():
    name = request.args.get('page')
    index = Linked_name.index(name)
    name = data.page('pre',index)
    return redirect(url_for('movie_name',name = name) )
    
   

@app.route('/crowdengine/next/')
def next():
    name = request.args.get('page')
    print(name)
    index = Linked_name.index(name)
    name = data.page('next',index)
    return redirect(url_for('movie_name',name = name))

@app.route('/crowdengine/write/',methods=["POST"])   
def write_db():
    
    if request.method == 'POST':
        id={
'Sanjana':'sahj@gmail.com' ,'Aswin':'aswin@gmail.com','praveena':'praveena@gmail.com'
} 
    

        movie = request.form["movie"]
        actor = request.form["name"]
        dur = request.form["Duration"]
        hair = request.form["Hairstyle"]
        role = request.form["Role"]
        hit = request.form["Hit"]
        color = request.form["Dresscolor"]

        data.write(movie=movie,actor=actor,duration=dur,hairstyle=hair,role=role,dresscolor=color,target=hit)

    return render_template('movies.html',p="ok" ,name=movie)

@app.route('/download/')
def download():
    return data.get_csv(a = app)

if __name__ == "__main__":
    app.run()
    # a=get_xl_data()

    
