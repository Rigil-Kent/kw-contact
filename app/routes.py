from app import app, db
import sqlite3
import pandas as pd
import xlsxwriter, xlrd, csv, glob
from openpyxl import load_workbook
from datetime import datetime
from app.forms import ContactForm
from app.models import Contact
from flask import render_template, redirect, flash, Markup, url_for, request


@app.route('/', methods=['GET', 'POST'])
def index():
    form = ContactForm()

    if form.validate_on_submit():
        contact = Contact()
        contact.first_name = form.first_name.data
        contact.last_name = form.last_name.data
        contact.middle_name = form.middle_name.data
        contact.prefix = form.prefix.data
        contact.suffix = form.suffix.data
        contact.legal_name = contact.first_name + " " + contact.middle_name + " " + contact.last_name
        contact.about = form.about.data
        contact.country_code = form.country_code.data
        contact.home_phone = form.home_phone.data
        contact.work_phone = form.work_phone.data
        contact.other_phone = form.other_phone.data
        contact.email = form.email.data
        contact.email2 = form.email2.data
        contact.street1 = form.street1.data
        contact.city1 = form.city1.data
        contact.postal_code1 = form.postal_code1.data
        contact.state1 = form.state1.data
        contact.street2 = form.street2.data
        contact.city2 = form.city2.data
        contact.postal_code2 = form.postal_code2.data
        contact.state2 = form.state2.data
        contact.company_name = form.company_name.data
        contact.company_street = form.company_street.data
        contact.company_city = form.company_city.data
        contact.company_state = form.company_state.data
        contact.company_postal = form.company_postal.data
        contact.company_country = form.company_country.data
        contact.company_title = form.company_title.data
        contact.company_department = form.company_department.data
        contact.birthday = form.birthday.data
        contact.home_anniversary = form.birthday.data
        contact.agent = form.agent.data
        contact.vendor = form.vendor.data
        contact.loan_officer = form.loan_officer.data
        contact.talent = form.talent.data
        contact.builder = form.builder.data
        contact.reo = form.reo.data
        contact.short_sale = form.short_sale.data
        contact.expired = form.expired.data
        contact.fsbo = form.fsbo.data
        contact.sphere = form.sphere.data
        contact.allied_resource = form.allied_resource.data
        contact.referral_sources = form.referral_sources.data
        contact.first_time = form.first_time.data
        contact.tags = form.tags.data
        contact.notes = form.notes.data
        contact.facebook = form.facebook.data
        contact.twitter = form.twitter.data
        contact.linkedin = form.linkedin.data
        contact.google = form.google.data
        contact.pintrest = form.pintrest.data
        contact.instagram = form.instagram.data
        
        
        db.session.add(contact)
        db.session.commit()

        message = Markup('<div class="alert alert-success alert-dismissible"><button type="button" class="close" data-dismiss="alert">&times;</button>Contact {} added successfully.</div>'.format(contact.legal_name))
        flash(message)
        return redirect(url_for('contacts'))
    return render_template('index.html', form=form)



@app.route('/contacts')
def contacts():
    contacts = Contact.query.all()
    count = Contact.query.count()
    form = ContactForm()
    
    if count == 0:
        message = Markup('<div class="mt-3 alert alert-info alert-dismissible"><button type="button" class="close" data-dismiss="alert">&times;</button>You currently have <strong>0</strong> contacts.</div>')
        flash(message)

    fields = [field.label.text for field in ContactForm() if field.label if field.label.text != 'CSRF Token' or field.label.text != 'Add Contact']
    return render_template('contacts.html', contacts=contacts)

@app.route('/export')
def export():
    form = ContactForm()
    
    # grab all ContactForm field labels
    fields = [field.label.text for field in ContactForm()]

    workbook = xlsxwriter.Workbook(r'app/static/kw_db.xlsx')
    worksheet = workbook.add_worksheet()

    # write Header 1
    worksheet.write('R1', 'Address 1')
    worksheet.write('V1', 'Address 2')
    worksheet.write('AA1', 'Company Address')
    worksheet.write('AH1', 'Key Dates')
    worksheet.write('AJ1', 'STATUS GROUP')
    worksheet.write('AK1', 'System Tags (Select Y if applies)')
    worksheet.write('AY1', 'User Tags')
    worksheet.write('AZ1', 'Notes')
    worksheet.write('BD1', 'SOCIAL MEDIA')

    # write form fields to Header 2
    for i, d in enumerate(fields[:-3]):
        worksheet.write(1, i, d)

    # write db data to excel file
    row = 2
    col = 0
    contacts = Contact.query.all()
    for i, contact in enumerate(contacts):
        data = [contact.first_name, contact.middle_name, contact.last_name, contact.prefix, contact.suffix, contact.legal_name, contact.about, contact.country_code, contact.mobile_phone, contact.country_code, contact.home_phone, contact.country_code, contact.work_phone, contact.country_code, contact.other_phone, contact.email, contact.email2, contact.street1, contact.city1, contact.state1, contact.postal_code1, contact.street2, contact.city2, contact.state2, contact.postal_code2, contact.company_name, contact.company_street, contact.company_city, contact.company_state, contact.company_postal, contact.company_country, contact.company_department, contact.company_title, contact.birthday, contact.home_anniversary, "Null", contact.agent, contact.vendor, contact.investor, contact.loan_officer, contact.talent, contact.builder, contact.reo, contact.short_sale, contact.expired, contact.fsbo, contact.sphere, contact.allied_resource, contact.referral_sources, contact.first_time, contact.tags, contact.notes, contact.source, contact.other_source, contact.date_added, contact.facebook, contact.twitter, contact.linkedin, contact.google, contact.pintrest, contact.instagram]
        for j, d in enumerate(data):
            worksheet.write(row + i, col + j, d)

    workbook.close()

    # convert to csv for download
    df = pd.read_excel('app/static/kw_db.xlsx', sheetname='Sheet1', header=None)
    columns = df.iloc[:2]
    df = df.iloc[2:]
    columns = columns.fillna('')
    columns = pd.MultiIndex.from_arrays(columns.values.tolist())
    df.columns = columns
    df.to_csv('app/static/kw_db.csv', index=False, na_rep="")
    message = Markup('<div class="mt-2 alert alert-info"><h4><span class="fa fa-info-circle"></span> Here\'s your file</h4><a href="static/kw_db.csv">Download</a></div>')
    flash(message)

    return redirect(url_for('contacts'))


@app.route('/contacts/delete/<id>')
def delete_contact(id):
    contact = Contact.query.filter_by(id=id).first_or_404()
    db.session.delete(contact)
    db.session.commit()
    message = Markup('<div class="alert alert-success alert-dismissible"><button type="button" class="close" data-dismiss="alert">&times;</button>Contact {} deleted.</div>'.format(contact.first_name + " " + contact.last_name))
    flash(message)
    return redirect(url_for('contacts'))

@app.route('/contacts/view/<id>')
def view_contact(id):
    contact = Contact.query.filter_by(id=id).first_or_404()
    return render_template('contact.html', contact=contact)

@app.route('/contacts/edit/<id>')
def edit_contact(id):
    contact = Contact.query.filter_by(id=id).first_or_404()
    form = ContactForm()
    return render_template('edit_contact.html', contact=contact, form=form)