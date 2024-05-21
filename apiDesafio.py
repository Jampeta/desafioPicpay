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
    value = data.get('value')
    id_payer = data.get('payer')
    id_payee = data.get('payee')
 
    if not value or not id_payer or not id_payee:
        return jsonify({"Erro": "Valor, pagador e recebedor são obrigatórios!"}), 400
 
    try:
        payer = collection.find_one({'id': id_payer})
        payee = collection.find_one({'id': id_payee})
 
        if payer is None or payee is None:
            return jsonify({"Erro": "Pagador ou recebedor não encontrados"}), 404
 
        if payer['tipo'] == 'lojista':
            return jsonify({'Erro': 'O Usuário não pode fazer transferencia!'}), 403
 
        saldo_payer = float(payer['saldo'])
        if saldo_payer < value:
            return jsonify({"Erro": "Saldo Insuficiente!"}), 400
 
        new_saldo_payer = saldo_payer - value
        collection.update_one({'id': id_payer}, {'$set': {'saldo': new_saldo_payer}})
 
        saldo_payee = float(payee['saldo'])
        new_saldo_payee = saldo_payee + value
        collection.update_one({'id': id_payee}, {'$set': {'saldo': new_saldo_payee}})

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
        <h2>Resumo da transferencia</h2>
        <hr></hr>

        <p>O valor transferido foi <strong>R${value}</strong></p>
        <p>Nome do remetente: {nome_payer}</p>
        <p>Nome do destinatario: {nome_payee}</p>
    
        <p>Abs,</p>
        <p>Picpay de python :)</p>
        """

        email.Send()
        
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


if __name__ == "__main__":
    app.run(debug=True, port=5000, host="0.0.0.0")
