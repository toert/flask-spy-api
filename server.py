from flask import Flask, render_template, request, make_response, send_file
from spy_words import *


app = Flask(__name__)


@app.route('/load', methods=['GET'])
def load():
    return send_file('result.xlsx')


@app.route('/', methods=['GET'])
def form():
    return render_template('form.html', text=('', ''), last_time=LAST_UPDATE_TIME)

@app.route('/', methods=['POST'])
def typograf():
    login = request.form['login']
    password = request.form['token']
    limit = request.form['limit']
    input_text = request.form['text']
    words = [word for word in input_text.split('\r\n')]
    if login and password and input_text:
        output_text = parse_info(words, login, password, limit)
        #return render_template('form.html', text=(input_text, output_text), token=password, login=login, limit=limit)
        #response = make_response(output_text)
        # This is the key: Set the right header for the response
        # to be downloaded, instead of just printed on the browser
        #response.headers["Content-Disposition"] = "attachment; filename=result.csv"
        global LAST_UPDATE_TIME
        LAST_UPDATE_TIME = datetime.strftime(datetime.now(timezone('Europe/Moscow')), '%H:%M:%S')
        return send_file(output_text)
    else:
        output_text = 'Неправильный пароль!'
        return render_template('form.html', text=(input_text, output_text), token=password, login=login, limit=limit,
                last_time=LAST_UPDATE_TIME)


if __name__ == "__main__":
    app.run(debug=True)