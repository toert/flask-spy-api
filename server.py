from flask import Flask, render_template, request
from spy_words import *


app = Flask(__name__)


@app.route('/', methods=['GET'])
def form():
    return render_template('form.html', text=('', ''))


@app.route('/', methods=['POST'])
def typograf():
    login = request.form['login']
    password = request.form['token']
    limit = request.form['limit']
    input_text = request.form['text']
    words = [word for word in input_text.split('\r\n')]
    if login and password and input_text:
        output_text = parse_info(words, login, password, limit)
        return render_template('form.html', text=(input_text, output_text), token=password, login=login, limit=limit)
    else:
        output_text = 'Неправильный пароль!'
        return render_template('form.html', text=(input_text, output_text), token=password, login=login, limit=limit)


if __name__ == "__main__":
    app.run(debug=True)