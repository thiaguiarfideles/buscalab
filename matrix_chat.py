import os
import asyncio
from flask import Flask, request, jsonify, render_template
import pandas as pd

app = Flask(__name__)

async def read_excel(file_name):
    loop = asyncio.get_running_loop()
    df = await loop.run_in_executor(None, pd.read_excel, file_name)
    return df

async def get_dataframe():
    tasks = [
        read_excel('exa_lab.xlsx'),
        read_excel('exa_lab_df.xlsx'),
        read_excel('exa_lab_nordest.xlsx')
    ]
    df_sp, df_df, df_nd = await asyncio.gather(*tasks)
    df = pd.concat([df_sp, df_df, df_nd], axis=0)
    return df

async def get_response(message):
    # Carrega a planilha em um DataFrame
    df = await get_dataframe()

    if message.isdigit():
        # Pesquisa o nome do exame pelo código do exame
        exame = df.loc[df['NR_SEQ_EXAME'] == int(message), ['NM_EXAME', 'RESPOSTA']].values
        if len(exame) > 0:
            return [f"Dados do Exame: {exame[0]} - {exame[1]}" for exame in exame]
        else:
            return ["Não foi encontrado nenhum exame com este código."]
    else:
        # Pesquisa o nome do exame que faz referência a NM_EXAME ou a RESPOSTA
        exames = df.loc[
            (df['NM_EXAME'].str.contains(message, case=False)) |
            (df['RESPOSTA'].str.contains(message, case=False)),
            ['NM_EXAME', 'RESPOSTA']
        ].values

        if len(exames) > 0:
            return [f"Dados do Exame: {exame[0]} - {exame[1]}" for exame in exames]
        else:
            return ["Não foi encontrado nenhum exame com este nome ou resposta."]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/chatbot', methods=['POST'])
async def process_message():
    message = request.form['message']
    response = await get_response(message)
    if len(response) == 0:
        flag = False
    else:
        flag = True
    return render_template('chatbot.html', message=response, flag=flag)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
