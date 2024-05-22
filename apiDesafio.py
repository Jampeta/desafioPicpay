import json
import os
from pymongo import MongoClient
from flask import Flask, Response, request, jsonify
import win32com.client as win32
import pythoncom

app = Flask(__name__)

client = MongoClient("mongodb://localhost:27017/")
db = client["picpay"]
collection = db["contas"] 


@app.route("/")
def base():
    return Response(
        response=json.dumps({"Status": "UP"}), status=200, mimetype="application/json"
    )

@app.route("/mostrar", methods=["GET"])
def mongo_read():
    
    documents = collection.find()
    output = [
            {item: data[item] for item in data if item != "_id"} for data in documents
        ]
    return Response(
        response=json.dumps(output), status=200, mimetype="application/json"
    )

@app.route('/add', methods=['POST'])
def adicionar():
    dados = request.get_json()

    campos_obrigatorios = ['id', 'nome', 'cpf', 'email', 'senha', 'saldo', 'tipo']
    for campo in campos_obrigatorios:
        if campo not in dados or not dados[campo]:
            return jsonify({'erro': f'O campo {campo} é obrigatório e não pode estar vazio.'}), 40
        
    cod = dados.get('id')
    objeto = collection.find_one({'id': cod})

    if objeto:
        return jsonify({'mensagem': 'Documento já existe'}), 400

    try:
        resultado = collection.insert_one(dados)
        return jsonify({'mensagem': 'Documento adicionado com sucesso!', 'id': str(resultado.inserted_id)}), 201
    except Exception as e:
        return jsonify({'erro': str(e)}), 500



@app.route("/transferir", methods=["PUT"])
def transfer():
    data = request.json
    valor = data.get('valor')
    id_remetente = data.get('remetente')
    id_destinatario = data.get('destinatario')
 
    if not valor or not id_remetente or not id_destinatario:
        return jsonify({"Erro": "Valor, pagador e recebedor são obrigatórios!"}), 400
 
    try:
        remetente = collection.find_one({'id': id_remetente})
        destinatario = collection.find_one({'id': id_destinatario})
        
 
        if remetente is None or destinatario is None:
            return jsonify({"Erro": "Pagador ou recebedor não encontrados"}), 404
 
        if remetente['tipo'] == 'lojista':
            return jsonify({'Erro': 'O Usuário não pode fazer transferencia!'}), 403
 
        saldo_remetente = float(remetente['saldo'])
        if saldo_remetente < valor:
            return jsonify({"Erro": "Saldo Insuficiente!"}), 400
        
        saldo_temp_remetente = remetente['saldo']
        try:
            new_saldo_remetente = saldo_remetente - valor
            collection.update_one({'id': id_remetente}, {'$set': {'saldo': new_saldo_remetente}})
        except Exception as e:
            collection.update_one({'id': id_remetente}, {'$set': {'saldo': saldo_temp_remetente}})
            return jsonify({"Erro": str(e)}), 500
        
        saldo_temp_destinatario = destinatario['saldo']
        try:
            saldo_destinatario = float(destinatario['saldo'])
            new_saldo_destinatario = saldo_destinatario + valor
            collection.update_one({'id': id_destinatario}, {'$set': {'saldo': new_saldo_destinatario}})
        except Exception as e:
            collection.update_one({'id': id_destinatario}, {'$set': {'saldo': saldo_temp_destinatario}})
            return jsonify({"Erro": str(e)}), 500

        sendEmail(remetente, destinatario, valor)
        
        return jsonify({'Status': 'Transferencia Concluida!'}), 200
    except Exception as e:
        return jsonify({"Erro": str(e)}), 500
    

@app.route("/delete", methods=["DELETE"])
def delete():
    data = request.json
    cod = data.get('id');
    document = collection.find_one({'id': cod})

    if document:
        try:
            collection.delete_one({'id': cod})
            return jsonify({"Message": "Usuario deletado"}), 400
        except Exception as e:
            return jsonify({'Erro': str(e)}), 500
    else:
        return jsonify({"Error": "O usuario não existe"})

def sendEmail(payer, payee, value):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    nome_payer = payer.get('nome')
    nome_payee = payee.get('nome')
    email_payer = str(payer.get('email'))
    email_payee = str(payee.get('email'))

    email.To = email_payee + "; " + email_payer
    email.Subject = "Transferencia ocorrida PICPAY"
    email.HTMLBody = f"""
    <center><h2>Resumo da transferencia</h2></center>
    <hr></hr>

    <center><h1 style="color: #00A000"><strong>R${value}</strong></h1></center>
    <center><div style="display: inline-block">
        <p>Nome do remetente: <strong>{nome_payer}</strong></p>
        <p>Nome do destinatario: <strong>{nome_payee}</strong></p>
    </div></center>
    <hr></hr>
    <p><strong>Ass:</strong> Picpay de python ❇️</p>
    """

    email.Send()

if __name__ == "__main__":
    app.run(debug=True, port=5000, host="0.0.0.0")
