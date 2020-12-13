from flask import Flask, render_template, request, redirect, url_for, session, Response, make_response, send_file
import pandas as pd
import MySQLdb
import pandas.io.sql as psql
import re
import json
import io
from io import BytesIO
from wtforms import TextField, Form, SelectField, RadioField
from wtforms.validators import InputRequired
from flask_bootstrap import Bootstrap
from flask_login import LoginManager
from werkzeug.security import generate_password_hash, check_password_hash
import numpy as np
from datetime import datetime

app = Flask(__name__)

app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

class OrderForm(Form):
    sales_order = TextField(render_kw = {"placeholder" : "Insert Sales Order ID" }, id='sales_order')
    channel_order = TextField(render_kw = {"placeholder" : "Insert Channel Order ID" }, id='channel_order')
    shipping_cost = TextField(render_kw = {"placeholder" : "Insert Shipping Cost" }, id='shipping_cost')
    customer_name = TextField(render_kw = {"placeholder" : "Insert Customer Name" }, id='customer_name')
    shipping_address = TextField(render_kw = {"placeholder" : "Insert Shipping Address" }, id='shipping_address')
    shipping_phone = TextField(render_kw = {"placeholder" : "Insert Shipping Phone" }, id='shipping_phone')
    courier_name = TextField(render_kw = {"placeholder" : "Insert Courier Name" }, id='courier_name')
    courier_phone = TextField(render_kw = {"placeholder" : "Insert Courier Phone" }, id='courier_phone')

class ProductForm(Form):
    product_barcode = TextField(render_kw = {"placeholder" : "Insert Product Barcode" }, id='product_barcode')

@app.route('/', methods = ['GET', 'POST'])
def main():
    if 'user' in session:
        return redirect(url_for('create'))
    if request.method == 'POST' and 'username' in request.form and 'password' in request.form:
        username = request.form['username']
        password = request.form['password']

        if username == 'ecommerce' and password == 'ecommerce':
            session['user'] = True
            session['messages'] = 'Loggin Successfully'
            return redirect(url_for('create'))
        else :
            msg = 'Incorrect username or password'
            return render_template('home.html', msg = msg)    
    return render_template('home.html')

@app.route('/Create', methods = ['GET', 'POST'])
def create():
    if 'user' not in session:
        return redirect(url_for('main'))
    form = OrderForm(request.form)
    if request.method == 'POST':
        result = request.form
        session['Create'] = result
        product = pd.DataFrame(columns = ['No', 'SKU', 'Item Name', 'Item Price', 'Quantity', 'Total'])
        session['Product'] = product.to_dict('list')
        return redirect(url_for('insert_product'))
    return render_template('create.html', form = form)


@app.route('/Insert', methods = ['GET', 'POST'])
def insert_product():
    if 'user' not in session:
        return redirect(url_for('main'))
    db_product = pd.read_excel(r'Data utk POS.xlsx', sheet_name='Daftar Produk')
    form = ProductForm(request.form)
    if 'insert' in request.form:
        if request.method == 'POST':
            result = request.form
            product = session.pop('Product')
            product = pd.DataFrame(product)
            product_barcode = str(result['product_barcode'])
            indeks = db_product[db_product['Barcode'].astype(str) == product_barcode].index[0]
            product_name = db_product['Item Name'][indeks]
            product_sku = db_product['SKU'][indeks]
            product_price = db_product['price after disc (harga coret)'][indeks]
            if product_name not in product['Item Name'].values:
                row = pd.DataFrame([[product_sku, product_name, product_price, 1, product_price]], columns = ['SKU', 'Item Name', 'Item Price', 'Quantity', 'Total'])
                product = product.append(row, ignore_index = True, sort = False)
            else :
                indeks = product[product['Item Name'] == product_name].index[0]
                product['Quantity'][indeks] = product['Quantity'][indeks] + 1
                product['Total'][indeks] = product['Total'][indeks] + product['Item Price'][indeks]
            
            print(product)
            product = product[['SKU', 'Item Name', 'Item Price', 'Quantity', 'Total']]
            session['Product'] = product.to_dict('list')
            display_product = product.copy()
            for i in display_product.columns:
                if i == 'SKU' or i == 'Item Name':
                    display_product[i] = display_product[i].astype(str).str.replace('.0', '', regex = False)
                else :
                    display_product[i] = display_product[i].astype(int).apply(lambda x: f'{x:,}')
            return render_template('insert.html', form = form, data = [display_product.to_html(classes = 'data')], titles = display_product.columns.values)
    
    elif 'reset' in request.form:
        if request.method == 'POST':
            product = pd.DataFrame(columns = ['SKU', 'Item Name', 'Item Price', 'Quantity', 'Total'])
            session['Product'] = product.to_dict('list')
            display_product = product.copy()
            for i in display_product.columns:
                if i == 'SKU' or i == 'Item Name':
                    display_product[i] = display_product[i].astype(str).str.replace('.0', '', regex = False)
                else :
                    display_product[i] = display_product[i].astype(int)
            return render_template('insert.html', form = form, data = [display_product.to_html(classes = 'data')], titles = display_product.columns.values)

    product = pd.DataFrame(columns = ['SKU', 'Item Name', 'Item Price', 'Quantity', 'Total'])
    session['Product'] = product.to_dict('list')
    return render_template('insert.html', form = form)

@app.route('/Verif', methods = ['GET', 'POST'])
def verif():
    order_info = session.pop('Create')
    session['Create'] = order_info
    product = session.pop('Product')
    session['Product'] = product

    print(order_info)
    order_info = pd.DataFrame(order_info, index = [0])
    product = pd.DataFrame(product)

    display_product = product.copy()
    for i in display_product.columns:
        if i == 'SKU' or i == 'Item Name':
            display_product[i] = display_product[i].astype(str).str.replace('.0', '', regex = False)
        else :
            display_product[i] = display_product[i].astype(int)

    if request.method == 'POST':
        order_info['tmp'] = 1
        product['tmp'] = 1
        order_product = product.merge(order_info, how = 'left', on = 'tmp')
        order_product = order_product.drop(['tmp'], axis = 1)
        data_all = pd.DataFrame()
        
        for index, row in order_product.iterrows():
            order_date = pd.to_datetime(datetime.today())
            channel = 'Grab Mart'
            sales_order = row['sales_order']
            channel_order = row['channel_order']
            invoice_number = channel_order
            customer_name = row['customer_name']
            item_name = row['Item Name']
            sku = row['SKU']
            quantity = row['Quantity']
            price = row['Item Price']
            shipping_cost = row['shipping_cost']
            shipping_name = customer_name
            shipping_address1 = row['shipping_address']
            shipping_address2 = np.nan
            shipping_city = np.nan
            shipping_zip = np.nan
            shipping_province= np.nan
            shipping_country = np.nan
            shipping_phone = row['shipping_phone']
            shipping_courier = 'Grab Express'
            awb = np.nan
            row_append = pd.DataFrame([[order_date, channel, sales_order, channel_order, invoice_number, customer_name, item_name, sku, quantity, price,shipping_cost,shipping_name, shipping_address1, shipping_address2,shipping_city, shipping_zip, shipping_province, shipping_country, shipping_phone, shipping_courier, awb]], columns = ['Order date', 'Channel', 'Sales Order ID', 'Channel Order ID', 'Invoice Number', 'Customer Name', 'Item Name', 'SKU', 'Quantity', 'Price', 'Shipping Cost', 'Shipping Name', 'Shippign Address1', 'Shipping Address2', 'Shipping City', 'Shipping Zip', 'Shipping Province', 'Shipping Country', 'Shipping Phone', 'Shipping Courier', 'AWB'])
            data_all = data_all.append(row_append, ignore_index = True, sort = False)
        
        print(data_all)
        session['All'] = data_all.to_dict('list')
        session['msg'] = 'Submit Data Successful'
        return redirect(url_for('main'))
    
    return render_template('verif.html', data = [order_info.to_html(classes = 'data')], titles = order_info.columns.values, data2 = [display_product.to_html(classes = 'data')], titles2 = display_product.columns.values)

@app.route('/Download', methods = ['GET', 'POST'])
def download():
    data_all = session.pop('All')
    session['All'] = data_all
    
    data_all = pd.DataFrame(data_all)
    data_all = data_all[['Order date', 'Channel', 'Sales Order ID', 'Channel Order ID', 'Invoice Number', 'Customer Name', 'Item Name', 'SKU', 'Quantity', 'Price', 'Shipping Cost', 'Shipping Name', 'Shippign Address1', 'Shipping Address2', 'Shipping City', 'Shipping Zip', 'Shipping Province', 'Shipping Country', 'Shipping Phone', 'Shipping Courier', 'AWB']]
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    data_all.to_excel(writer, sheet_name='Sheet1')
    writer.save()
    output.seek(0)
    return send_file(output, attachment_filename='output.xlsx', as_attachment=True)






