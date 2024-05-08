from EstacaZero import wordExport, plotagemAoki, plotagemDQ, plotCompararAokieDecourt, resultadosAoki, resultadosDQ, comparativoAokieDecourt
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

@app.route('/')

def estacaZero():

    return render_template('main.html')

"""
@app.rout('/tabelas', methods=['POST'])

def tabelas():
        
        resultadosAoki(listaSolos, tipoEstaca, diametroEstaca, listaNspt) #Vai retornar o Dataframe de Aoki
        resultadosDQ(listaSolos, tipoEstaca, diametroEstaca, listaNspt) #Vai retornar o Dataframe do Decourt Quaresma
        comparativoAokieDecourt(listaSolos, tipoEstaca, diametroEstaca, listaNspt) #Vai retornar o Dataframe do Comparativo entre os métodos
     
        return
"""

"""
@app.route('/graficos', methods=['POST'])

def plotar():
        
        plotagemAoki(listaSolos, tipoEstaca, diametroEstaca, listaNspt, cargaAdmissivel, niveldAgua) #Vai retorar a imagem do gráfico de Aoki Velloso em bytes

        plotagemDQ(listaSolos, tipoEstaca, diametroEstaca, listaNspt, cargaAdmissivel, niveldAgua)  #Vai retorar a imagem do gráfico de Decourt Quaresma em bytes

        plotCompararAokieDecourt(listaSolos, tipoEstaca, diametroEstaca, listaNspt, cargaAdmissivel, niveldAgua) #Vai retorar a imagem do gráfico do Comparativo em bytes
        
        return 
"""

@app.route('/retornar', methods=['POST'])

def retornar():

    return  render_template('main.html') #Criei, para tentar implementar um botão que retorna para a página main.html na página resultados.html 


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