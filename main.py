from flask import Flask, redirect, url_for, request ,render_template, send_file
import xlwings as xw
app = Flask(__name__,template_folder='.') 
  
@app.route('/') 
def input_screen(): 
   return render_template('landing.html')
  
@app.route('/vbachange',methods = ['POST', 'GET']) 
def vbachange():
    wb = xw.Book('VBAFIle.xlsm') 
    sht = wb.sheets['Sheet1']
    if request.method == 'POST':
    	user = request.form['nm'] 
    	sht.range('A1').value = user
    	output = wb.macro('Output')
    	output()
    	wb.save()
    	return send_file('VBAFIle.xlsm',as_attachment=True, attachment_filename='VBAFIle.xlsm')
    else: 
      	user = request.form['nm'] 
      	sht.range('A1').value = user
      	output = wb.macro('Output')
      	output()
      	wb.save()
      	return send_file('VBAFIle.xlsm',as_attachment=True, attachment_filename='VBAFIle.xlsm')

if __name__ == '__main__': 
   app.run(debug = True) 