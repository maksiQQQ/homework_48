from openpyxl import load_workbook

from flask import Flask, render_template, request

app = Flask(__name__)
excel = load_workbook('goods.xlsx')


@app.route('/')
def main():
    page = excel['Лист1']
    row = page['A']
    goods_list = []
    for i in range(len(row)):
        goods_list.append(row[i].value)
    if len(row) > len(goods_list):
        goods_list.append(page[f'A{len(goods_list)+1}'].value)
    return render_template('index.html', goods=goods_list)


@app.route('/add/', methods=['POST'])
def add():
    good = request.form['good']
    page = excel['Лист1']
    row = page['A']
    goods_list = []
    for i in range(len(row)):
        goods_list.append(row[i].value)
    if len(row) > len(goods_list):
        goods_list.append(page[f'A{len(goods_list)+1}'].value)
    page[f'A{len(goods_list)+1}'] = good
    excel.save('goods.xlsx')
    return render_template('add.html')


if __name__ == '__main__':
    app.run(debug=True)