import pywhatkit
# syntax: phone number with country code, message, hour and minutes
import openpyxl as xl
wb=xl.load_workbook(r'E:\2022-23 Even\ISPRC\SRIT HACKETHON\Email invitation\Book1_Project.xlsx')
sheet1=wb.get_sheet_by_name('Sheet1')
names=[]
emails=[]
team=[]
team_id=[]
cat=[]
mob=[]
for cell in sheet1['A']:
    emails.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

for cell in sheet1['C']:
    team.append(cell.value)

for cell in sheet1['D']:
    cat.append(cell.value)

for cell in sheet1['E']:
    team_id.append(cell.value)

for cell in sheet1['F']:
    mob.append(cell.value)


for i in range(len(emails)):
    message = ''' Hello {}, Greetings from SRIT ISPRC, Congratulations !!!  on being shortlisted for the SRIT HACKATHON 2023 .
           As per our records,
           Your Team Name : {},
           Track : {},
           and the generated Team ID is {}.
           You are cordially welcome to  the SRIT-Hackathon by Sri Ramakrishna Institute of Technology,
           Coimbatore on 13.02.2023 at 9:00 AM at SRIT. Kindly follow the list of Instructions given below

           ADDITIONAL INFORMATION
                1.	Registration
                    a)	Registration counter will be in the entrance of the college
                    b)	Students must need to bring their ID cards and show it in the registration desk
                    c)	Students must sign the attendance form in the registration Desk
                    d)	Timing : 13th February 9.00am-10.00am
                2.	Exhibition & Competition
                        13 February

                        a)	Project Exhibition, Poster Presentation , Paper Presentation : 10.30am – 3.00pm
                        b)	Project Exhibition and Poster Presentation is open to Internal faculty, Students and Public
                        c)	Judges will be visiting the respective place for evaluation (Project Exhibition, Poster Presentation). 
                        d)	Evaluation starts at 11.00 am
                        e)	At least one registered participant from the respective project or Poster group should occupy the place in the above mentioned timings.

                3.	Food
                        a)	No Food and Refreshments will be provided for the participants(Participants can utilize the canteen facility by payment)
                        b)	Digital payment facilities are available @ counter.

                4.Accommodation
                        a)	Participants who are located more than 300 Kms only, will be allowed to accommodate in Hostel.
                        b)	The Hostel accommodation will be provided in First Come and First Serve basics
                                (15 boys and 15 Girls only will be accommodated. Accommodation is chargeable.(For more details Mr R Immanual AP/MECH ,+9677817992)
                5.Important Instructions:
                        a)	Participants should bring their own requirements for project , Paper and Poster Presentation
                                (Laptop, Extension box, Pen drive and etc)
                        b)	No flammable object are allowed inside the campus of SRIT
                        c)	Participants are instructed to maintain a disciplined behavior throughout the Competition 
                        d)	Only Single Phase Current will be provided to the Project 
                        e)	Any of the misbehavior activities will lead to termination of the team
                        f)	For project Track Must bring prototype/product
                        g)	For poster Track must bring the printout copy of the poster as per the given template  in A2 size
                                (Template download Link :http://shorturl.at/hkOT1)
                        h)	For paper presentation Track bring the soft copy of ppt as per the given template
                                (Template download Link : http://shorturl.at/dwO16)
                6.Certificates & Award ceremony
                        a.	Winners Certificate will be provided during the awarding ceremony(3.00pm to 5.00pm)
                        b.	Participation certificates will be provided after awarding Ceremony
                        c.	Venue-PG Seminar Hall

                7.Facilities provided in the place 
                        a.	One table
                        b.	If One plug point needed for Project Track kindly mail to immanual.me@srit.org with your Team ID
                        c.	If Wi-Fi is needed kindly mail to immanual.me@srit.org
                        d.      Kindly bring Extension cord.
                        e.     Plastic flex strictly not allowed to paste the posters as per SRIT regulation.

                Important Contacts:

                COORDINATOR	RAGHU NATH E	9965117100
                REGISTRATION	SATHYA J        8489333555 and BHUVANESHVARAN	9965360648
                ACCOMADATION	MR.IMMANUAL R	9677817992
                CERTIFICATES	MS.KIRTHIKA	9884711173





           
                                                                                                                          Warm Regards,
                                                                                                                          Mr R Immanual AP/MECH
                                                                                                                          Sri Ramakrishna Institute of Technology'''.format(names[i],team[i],cat[i],team_id[i])
    pywhatkit.sendwhatmsg_instantly('+91'+str(mob[i]), message, 15, tab_close=True)
   #pywhatkit.sendwhats_image('+91'+str(mob[i]), img_path="E:/2022-23 Even/ISPRC/ISPRC IDEATHON\SOCIAL MEDIA/INVITATION.jpeg", caption="Example image sent!")
