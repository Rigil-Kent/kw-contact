from app import app, db
import pandas as pd
from app.forms import ContactForm
from app.models import Contact
from flask import render_template, redirect, flash, Markup, url_for, request


@app.route('/', methods=['GET', 'POST'])
def index():
    form = ContactForm()
    writer = pd.ExcelWriter(url_for('static', filename='empty_db.xlsx'), engine='xlsxwriter')

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
        contact.company_street = form.company_street.data
        contact.company_city = form.company_city.data
        contact.company_state = form.company_state.data
        contact.company_postal = form.company_postal.data
        contact.company_country = form.company_country.data
        contact.company_title = form.company_title.data
        contact.company_department = form.company_department.data
        contact.birthday = form.birthday.data
        contact.home_anniversary = form.home_anniversary.data
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

        df = pd.DataFrame({
            'First Name': contact.first_name,
            'Middle Name': contact.middle_name,
            'Last Name': contact.last_name,
            'Prefix': contact.prefix,
            'Suffix': contact.suffix,
            'Full Legal Name': contact.legal_name,
            'About': contact.about,
            'Country Code': contact.country_code,
            'Mobile Phone': contact.other_phone,
            'Country Code': contact.country_code,
            'Home Phone': contact.home_phone,
            'Country Code': contact.country_code,
            'Work Phone': contact.work_phone,
            'Email': contact.email,
            'Email 2': contact.email2,
            'Street': contact.street1,
            'City': contact.city1,
            'State/Province': contact.state1,
            'Postal Code': contact.postal_code1,
            'Street': contact.street2,
            'City': contact.city2,
            'State/Province': contact.state2,
            'Postal Code': contact.postal_code2,
            'Company': contact.company_name,
            'Street': contact.company_street,
            'City': contact.company_city,
            'State/Province': contact.company_state,
            'Postal Code': contact.company_postal,
            'Country/Region': contact.company_country,
            'Title': contact.company_title,
            'Department': contact.company_department,
            'Birthday': contact.birthday,
            'Home Anniversary': contact.home_anniversary,
            'Required*': '',
            'Agent': contact.agent,
            'Vendor': contact.vendor,
            'Loan Officer': contact.loan_officer,
            'Talent': contact.talent,
            'Builder': contact.builder,
            'REO': contact.reo,
            'Short Sale': contact.short_sale,
            'Expired': contact.expired,
            'Sphere': contact.sphere,
            'Allied Resource': contact.allied_resource,
            'Referral Sources': contact.referral_sources,
            'First Time': contact.first_time,
            'Tags': contact.tags,
            'Notes': contact.notes,
            'Added': contact.date_added,
            'Facebook': contact.facebook,
            'Twitter': contact.twitter,
            'LinkedIn': contact.linkedin,
            'Google+': contact.google,
            'Pintrest': contact.pintrest,
            'Instagram': contact.instagram
        })

        db.session.add(contact)
        db.session.commit()

        df.to_excel(writer, sheet_name="Contacts")
        writer.save()
        message = Markup('<div class="alert alert-success alert-dismissible"><button type="button" class="close">&times;</button>Contact {} added successfully.</div>')
        flash(message)
        return redirect(url_for(request.url))
    return render_template('index.html', form=form)