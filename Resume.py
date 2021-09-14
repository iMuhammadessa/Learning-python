from docx import Document
from docx.shared import Inches

document = Document()

# Profile Picture
document.add_picture(
    'muhammad.png', width=Inches(2.0)
)

# Details (Name, Phone number, email)
name = input('What is your name? ')
phone_number = input('What is your phone_number? ')
email = input('What is your email? ')

document.add_paragraph(
    name + ' | ' + phone_number + ' | ' + email
)

#About me
document.add_heading('About me')
document.add_paragraph(
    input('Tell about yourself? ')
)

# Work experience
document.add_heading('Work Experience')
p = document.add_paragraph()

company = input('Enter company ')
from_date = input('From Date ')
to_date = input('To Date ')

p.add_run(company + ' ').bold = True
p.add_run(from_date + '-' + to_date + '\n').italic = True

experience_details = input(
    'Describe your experience at ' + company + ' ')
p.add_run(experience_details)

#more experiences
while True:
    has_more_experiences = input(
        'Do you have more experiences? Yes or No ')
    if has_more_experiences.lower() == 'yes':
        p = document.add_paragraph()

        company = input('Enter company ')
        from_date = input('From Date ')
        to_date = input('To Date ')

        p.add_run(company + ' ').bold = True
        p.add_run(from_date + '-' + to_date + '\n').italic = True

        experience_details = input(
            'Describe your experience at ' + company + ' ')
        p.add_run(experience_details)
    else:
        break

#more experiences
document.add_heading('Skills')
skill = input('Enter skill ')
p = document.add_paragraph(skill)
p.style = 'List Bullet'

while True:
    has_more_skills = input('Do you have more skills? YES or NO ')
    if has_more_skills.lower() == 'yes':
        skill = input('Enter skill ')
        p =document.add_paragraph(skill)
        p.style = 'List Bullet'
    else:
        break


document.save('cv.docx')