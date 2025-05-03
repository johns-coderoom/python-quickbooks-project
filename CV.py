from docx import Document
from docx.shared import Inches
import pyttsx3

def speak(text):
    pyttsx3.speak(text)

document = Document()


# profile picture
document.add_picture('me.jpg', width=Inches(2.0))

# name phone number email details
name = input('what is you name?')
speak('Hello' + name + 'how are you today?')
phone_number = input('what is your phone number?')
speak ( 'Your number is an mtn one with +233555731434')
email = input('what is your email address?') 
speak('Your email address is similar to your initials')


document.add_paragraph(name +' | ' + phone_number +' | '+ email)

# about me
document.add_heading('About_me')
about_me = input('Tell me about yourself?') 
document.add_paragraph(about_me) 

# work experiences 
document.add_heading('Work experience') 
p = document.add_paragraph()

company = input('Enter company')
From_date = input('from Date')
to_date = input('To date')

p.add_run(company + ' ' ).bold = True
p.add_run( From_date + '-' + to_date + '\n').italic

experience_details = input(
    ' Describe your experiences at' + company +' ')
p.add_run(experience_details)

# more experiences 
while True:
    has_more_experiences = input(
        'Do you have more responses? Yes or No')
    if has_more_experiences.lower() =='yes':
        p = document.add_paragraph()

        company = input('Enter company')
        From_date = input('from Date')
        to_date = input('To date')

        p.add_run(company + ' ' ).bold = True
        p.add_run( From_date + '-' + to_date + '\n').italic

        experience_details = input(
        ' Describe your experiences at' + company +' ')
        p.add_run(experience_details)
    else:
        break 

    # skills
    document.add_heading('Skills')
    skills = input('Enter skills')
    p=document.add_paragraph(skills)
    p.style='List Bullet'

    while True:
        has_more_skills = input('Do you have more skills? Yes or No')
        if has_more_skills.lower( ) =='yes':
            skills = input('Enter skills')
            p=document.add_paragraph(skills)
            p.style='List Bullet'
        else:
            break 

# footer
section = document.sections[0]
footer = section.footer
p = footer.paragraphs[0]
p.text = " CV generated  by the John using python programming language"

document.save('my.docx')