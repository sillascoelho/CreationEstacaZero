from EstacaZero import wordExport
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

@app.route('/')

def estacaZero():

    return render_template('main.html')

@app.route('/submitEstaca', methods=['POST', 'GET'])

def submitEstaca():

    if request.method == "POST":

        listaSolos = str(request.form.get('listaSolos'))
        listaSolos_str = listaSolos.split(",")
        listaSolos_int = [int(num.strip()) for num in listaSolos_str]

        listaNspt = str(request.form.get('listaNspt'))
        listaNspt_str = listaNspt.split(",")
        listaNspt_int = [int(num.strip()) for num in listaNspt_str]

        tipoEstaca = str(request.form.get('tipoEstaca'))
        diametroEstaca = float(request.form.get('diametroEstaca'))
        cargaAdmissivel = float(request.form.get('cargaAdmissivel'))
        niveldAgua = float(request.form.get('niveldAgua'))
        fileName = str(request.form.get('fileName'))       
       
        wordExport(listaSolos_int, tipoEstaca, diametroEstaca, listaNspt_int, cargaAdmissivel, niveldAgua, fileName)
        
        print(f"Recebido: {listaSolos_int}, {listaNspt_int}, {tipoEstaca}, {diametroEstaca}, {cargaAdmissivel}, {niveldAgua}")

        path_memoryWord = fileName + '.docx'

        return send_file(path_memoryWord, as_attachment=True, download_name=path_memoryWord)


if __name__ == '__main__':

    app.run(debug=True)