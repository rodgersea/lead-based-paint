from func_repo import *
import os
import subprocess
import pandas as pd
import traceback
from docx2pdf import convert
from PyPDF2 import PdfFileReader, PdfFileWriter
from PIL import Image
import fitz

pd.options.display.max_columns = None  # display options for table
pd.options.display.width = None  # allows 'print table' to fill output screen
pd.options.mode.chained_assignment = None  # disables error caused by chained dataframe iteration

# ----------------------------------------------------------------------------------------------------------------------
# global variables
insp_num = {'Chris Ciappina': '120303',
            'Fabrizzio Simoni': '120304',
            'Parker Alvis': '120301',
            'Larry Rockefeller': '120291',
            'Lee Clark': '120065',
            'Ryan Bumpass': '120310',
            'Rob Campbell': '120302',
            'Tom Majkowski': '120166',
            'Brian Long': 'unknown',
            'Kevin': 'unknown'}  # inspector numbers
name2sig = {'Chris Ciappina': 'chris_ciappina',
            'Fabrizzio Simoni': 'fabrizzio_simoni',
            'Parker Alvis': 'parker_alvis',
            'Larry Rockefeller': 'larry_rockefeller',
            'Lee Clark': 'lee_clark',
            'Ryan Bumpass': 'ryan_bumpass',
            'Rob Campbell': 'rob_campbell',
            'Tom Majkowski': 'tom_majkowski',
            'Brian Long': 'brian_long',
            'Kevin': 'unknown'}  # signature file name
table_names = ['Table 1: Lead-Based Paint¹',
               'Table 2: Deteriorated Lead-Based Paint¹',
               'Table 3: Lead Containing Materials²',
               'Table 4: Dust Wipe Sample Analysis',
               'Table 5: Soil Sample Analysis',
               'Table 6: Lead Hazard Control Options¹']
proj_num = '210289.00'

# ----------------------------------------------------------------------------------------------------------------------
# input: schedule as excel file
# output: variable df as dataframe containing pertinent information to the list of app numbers
# ----------------------------------------------------------------------------------------------------------------------
sched_lis = os.listdir('schedule_compile')
sched_hold = parse_excel('schedule_compile/' + str(sched_lis[0]))

arr = os.listdir('job_Folders')  # create array of file names
# create dataframe of rows pertaining to app numbers in "arr"
# create array of app numbers
arr1 = []
for x in range(len(arr)):
    store = arr[x][:9]
    arr1.append(store)

# create dataframe of info pertaining to each app number
df = pd.DataFrame(sched_hold.loc[sched_hold['APP'] == arr1[0]])

for y in range(1, len(arr1)):  # one by one concat remaining apps to df
    wkbk_hold = sched_hold.loc[sched_hold['APP'] == arr1[y]]
    df = pd.concat([df, wkbk_hold], axis=0)

df = df.reset_index(drop=True)  # reset indices of df
# ----------------------------------------------------------------------------------------------------------------------
# input: df1 as dataframe containing pertinent information to the list of app numbers
# output: LRA for each job folder present in folder "job_Folders/"
# output: full xrf data as excel file and pdf
# output: tables 1-5 as sheets in excel file
# ----------------------------------------------------------------------------------------------------------------------
subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"])  # kill excel at start

# call create_lra function on df1 to create LRA's for all apps in dataframe
for index, row in df.iterrows():
    thero = row.to_numpy()
    dispp('beholden', thero)

    app_lis = os.listdir('job_Folders/' + str(thero[0]) + ' - ' + str(thero[5]) + ' - ' + str(thero[6]) + '/' + str(thero[2]) + '_LBP')
    pdf_filename = os.listdir('job_Folders/' + str(thero[0]) + ' - ' + str(thero[5]) + ' - ' + str(thero[6]) + '/' + str(thero[2]) + '_LBP/lab_Results/')
    pdf_path1 = 'job_Folders/' + str(thero[0]) + ' - ' + str(thero[5]) + ' - ' + str(thero[6]) + '/' + str(thero[2]) + '_LBP/lab_Results/' + pdf_filename[0]
    gx = get_xrf(row.to_numpy())  # gx = raw xrf data
    xtab = xrf_tables(gx, pdf_path1)  # xtab =
    try:
        pathh = 'lead_Pit/LRA/finished_Docs/' + str(thero[0]) + '/' + str(thero[0]) + '_LRA.docx'
    except:
        traceback.print_exc()

    # save xrf_clean.xlsx in app folder
    save_xrf_clean_xlsx(gx,
                        thero)

    # create table 1: Lead Based Paint, save as table1_lbp.xlsx
    save_xrf_pos_xlsx(xtab,
                      thero)

    xrf_clean_excel2pdf(gx, thero)  # save clean excel file as pdf, use beholden to get .xlsx path name

    subprocess.call(["taskkill", "/f", "/im", "WINWORD.EXE"])  # kill excel at end

    create_lra(xtab,  # dflis
               thero,  # beholden
               insp_num,  # from global variables
               proj_num)  # from global variables

    path_lra = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_LRA.docx'
    convert(path_lra)

    create_lbpas(xtab,  # dflis
                 thero,  # beholden
                 insp_num,  # from global variables
                 name2sig)  # from global variables

    path_lbpas = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_LBPAS.docx'
    convert(path_lbpas)

    wavelis = ['form_5.0_Page_1',
               'form_5.0_Page_2',
               'form_5.1']
    wavepath = []
    for x in range(len(wavelis)):
        wavepath.append('job_Folders/' + thero[0] + ' - ' + thero[5] + ' - ' + thero[6] + '/' + thero[2] + '_LBP/' + wavelis[x] + '/' + os.listdir('job_Folders/' + thero[0] + ' - ' + thero[5] + ' - ' + thero[6] + '/' + thero[2] + '_LBP/' + wavelis[x])[0])
    for x in range(len(wavepath)):
        wavepath[x] = wavepath[x][:-4]

    for x in wavepath:
        img2pdf(x)

    floor_path = ('job_Folders/' + thero[0] + ' - ' + thero[5] + ' - ' + thero[6] + '/' + thero[2] + '_LBP/floorplan/' + os.listdir('job_Folders/' + thero[0] + ' - ' + thero[5] + ' - ' + thero[6] + '/' + thero[2] + '_LBP/floorplan')[0])[:-4]
    img2pdf(floor_path)

    input1 = fitz.open(pdf_path1)
    page_end = input1.loadPage(3)
    pix = page_end.get_pixmap()
    outputt = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.png'
    pix.save(outputt)

    imgpat = r'C:/Users/Elliott/pythonplay/lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.png'
    pdfpat = r'C:/Users/Elliott/pythonplay/lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.pdf'
    img1 = Image.open(imgpat)
    img2 = img1.convert('RGB')
    img2.save(pdfpat)

    resmain_reader = PdfFileReader(pdf_path1)
    resmain_writer = PdfFileWriter()
    for x in range(3):
        my_page = resmain_reader.getPage(x)
        resmain_writer.addPage(my_page)
    resmain_output = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_main.pdf'
    with open(resmain_output, 'wb') as output:
        resmain_writer.write(output)

    merge_lis = ['lead_Pit/LRA/finished_Docs/' + row.to_numpy()[0] + '/' + row.to_numpy()[0] + '_LRA.pdf',
                 'lead_Pit/reporting/LRA/attachments.pdf',
                 'lead_Pit/reporting/LRA/floor_Plan.pdf',
                 floor_path + '.pdf',  # place holder for actual floor plan
                 'lead_Pit/reporting/LRA/risk_Assessment.pdf',
                 wavepath[0] + '.pdf',
                 wavepath[1] + '.pdf',
                 wavepath[2] + '.pdf',
                 'lead_Pit/reporting/LRA/xrf_Photos.pdf',
                 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/xrf_clean.pdf',
                 'lead_Pit/reporting/LRA/xrf_Photos.pdf',  # place holder for xrf positive photo log

                 'lead_Pit/reporting/LRA/lab_Results.pdf',
                 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_main.pdf',
                 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.pdf',
                 'lead_Pit/reporting/LRA/method_all.pdf',
                 'lead_Pit/reporting/LRA/lbpas.pdf',

                 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_LBPAS.pdf',
                 'lead_Pit/reporting/LRA/xrf_all.pdf',
                 'lead_Pit/reporting/LRA/certs.pdf',
                 'lead_Pit/reporting/Licensure/Lead/' + thero[1] + '.pdf',
                 'lead_Pit/reporting/LRA/firm_license.pdf',

                 'lead_Pit/reporting/LRA/rebuild.pdf']

    merge_pdfs(merge_lis, 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_LBP_Report.pdf')

    subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"])  # kill excel at end
