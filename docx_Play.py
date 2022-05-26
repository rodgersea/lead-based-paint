
from docx2pdf import convert
from PyPDF2 import PdfFileReader, PdfFileWriter
from PIL import Image
from func_repo import *
from tools import *

import pandas as pd
import subprocess
import traceback
import fitz
import os

pd.options.display.max_columns = None  # display options for table
pd.options.display.width = None  # allows 'print table' to fill output screen
pd.options.mode.chained_assignment = None  # disables error caused by chained dataframe iteration

# ----------------------------------------------------------------------------------------------------------------------
# global variables
insp_num = {'Elliott Rodgers': '110341',
            'Chris Ciappina': '120303',
            'Fabrizzio Simoni': '120304',
            'Parker Alvis': '120301',
            'Larry Rockefeller': '120291',
            'Lee Clark': '120065',
            'Ryan Bumpass': '120310',
            'Rob Campbell': '120302',
            'Tom Majkowski': '120166',
            'Brian Long': 'unknown',
            'Kevin': 'unknown'}  # inspector numbers
name2sig = {'Elliott Rodgers': 'elliott_rodgers',
            'Chris Ciappina': 'chris_ciappina',
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
proj_num = '220083.00'
subprocess.call(["taskkill", "/f", "/im", "WINWORD.EXE"])  # kill word
cwd = os.path.abspath(os.path.dirname(__file__))

# ----------------------------------------------------------------------------------------------------------------------
# input: schedule as excel file
# output: variable df as dataframe containing pertinent information to the list of app numbers
# ----------------------------------------------------------------------------------------------------------------------
sched_lis = os.listdir(os.path.join(cwd, 'schedule_compile'))
sched_hold = parse_excel(os.path.join(cwd, 'schedule_compile', str(sched_lis[0])))

arr = os.listdir(os.path.join(cwd, 'job_Folders'))  # create array of file names
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
    full_app = thero[0] + ' - ' + thero[5] + ' - ' + thero[6]
    lead_str = thero[2] + '_LBP'
    app_data_pat = os.path.join(cwd, 'job_Folders', full_app, lead_str, 'app_Data')
    app_report_pat = os.path.join(cwd, 'finished_Docs', thero[0])

    fin_check_path = os.path.join(cwd, 'finished_Docs', str(thero[0]), thero[0] + '_LBP_Report_' + thero[11].strftime('%m%d%y') + '.pdf')
    if not os.path.exists(fin_check_path):
        dispp('beholden', thero)

        pdf_path1 = os.path.join(app_data_pat, 'lab_Results', os.listdir(os.path.join(app_data_pat, 'lab_Results'))[0])
        gx = get_xrf(row.to_numpy())  # gx = raw xrf data

        pb_res = pdf_scrape(pdf_path1)
        xtab = xrf_tables(gx, pb_res)
        xtab1 = xrf_tables(gx, pb_res)
        xtab2 = xrf_tables(gx, pb_res)
        try:
            pathh = os.path.join(app_report_pat, str(thero[0]) + '_LRA.docx')
        except:
            traceback.print_exc()

        # save xrf_clean.xlsx in app folder
        save_xrf_clean_xlsx(gx,
                            thero)

        # create table 1: Lead Based Paint, save as table1_lbp.xlsx
        save_xrf_pos_xlsx(xtab,
                          thero)

        xrf_clean_excel2pdf(gx, thero)  # save clean excel file as pdf, use beholden to get .xlsx path name

        subprocess.call(["taskkill", "/f", "/im", "WINWORD.EXE"])  # kill word

        create_photo_log(thero, cwd)

        create_lra(xtab1,  # dflis
                   thero,  # beholden
                   insp_num,  # from global variables
                   proj_num)  # from global variables

        path_lra = os.path.join(app_report_pat, thero[0] + '_LRA.docx')

        lra_pdf_ex = os.path.join(app_report_pat, thero[0] + '_LRA.pdf')
        if not os.path.exists(lra_pdf_ex):
            convert(path_lra)

        create_lbpas(xtab2,  # dflis
                     thero,  # beholden
                     insp_num,  # from global variables
                     name2sig)  # from global variables

        path_lbpas = os.path.join(app_report_pat, thero[0] + '_LBPAS.docx')

        lbpas_pdf_ex = os.path.join(app_report_pat, thero[0] + '_LBPAS.pdf')
        if not os.path.exists(lbpas_pdf_ex):
            convert(path_lbpas)

        wavelis = ['form_5.0_Page_1',
                   'form_5.0_Page_2',
                   'form_5.1']
        wavepath = []
        for x in range(len(wavelis)):
            wavey = os.listdir(os.path.join(app_data_pat, wavelis[x]))
            for y in wavey:
                if y[-4:] != '.pdf':
                    img2pdf(y)
                    wavepath.append(os.path.join(app_data_pat, wavelis[x], y[:-4] + '.pdf'))
                    break
                else:
                    wavepath.append(os.path.join(app_data_pat, wavelis[x], y))

        for x in os.listdir(os.path.join(app_data_pat, 'floorplan')):
            if x[-4:] != '.pdf':
                img2pdf(x)
                floor_path = os.path.join(app_data_pat, 'floorplan', x[:-4] + '.pdf')
                break
            else:
                floor_path = os.path.join(app_data_pat, 'floorplan', x)

        input1 = fitz.open(pdf_path1)
        page_end = input1.loadPage(3)
        pix = page_end.get_pixmap()
        outputt = os.path.join(cwd, 'finished_Docs', thero[0], 'res_end.png')
        pix.save(outputt)

        imgpat = os.path.join(cwd, 'finished_Docs', thero[0], 'res_end.png')
        pdfpat = os.path.join(cwd, 'finished_Docs', thero[0], 'res_end.pdf')

        img1 = Image.open(imgpat)
        img2 = img1.convert('RGB')
        img2.save(pdfpat)

        resmain_reader = PdfFileReader(pdf_path1)
        resmain_writer = PdfFileWriter()
        for x in range(3):
            my_page = resmain_reader.getPage(x)
            resmain_writer.addPage(my_page)
        resmain_output = 'finished_Docs/' + thero[0] + '/res_main.pdf'
        with open(resmain_output, 'wb') as output:
            resmain_writer.write(output)

        merge_lis = [os.path.join('finished_Docs', row.to_numpy()[0], row.to_numpy()[0] + '_LRA.pdf'),
                     'reporting_Docs/LRA/attachments.pdf',
                     'reporting_Docs/LRA/floor_Plan.pdf',
                     floor_path,
                     'reporting_Docs/LRA/risk_Assessment.pdf',
                     wavepath[0],
                     wavepath[1],
                     wavepath[2],
                     'reporting_Docs/LRA/xrf_Photos.pdf',
                     os.path.join('finished_Docs', thero[0], 'xrf_clean.pdf'),
                     os.path.join('finished_Docs', thero[0], thero[0] + '_photo_Log.pdf'),

                     'reporting_Docs/LRA/lab_Results.pdf',
                     os.path.join('finished_Docs', thero[0], 'res_main.pdf'),
                     os.path.join('finished_Docs', thero[0], 'res_end.pdf'),
                     'reporting_Docs/LRA/method_all.pdf',
                     'reporting_Docs/LRA/lbpas.pdf',

                     os.path.join('finished_Docs', thero[0], thero[0] + '_LBPAS.pdf'),
                     'reporting_Docs/LRA/xrf_all.pdf',
                     'reporting_Docs/LRA/certs.pdf',
                     os.path.join('reporting_Docs/Licensure/Lead', thero[1] + '.pdf'),
                     'reporting_Docs/LRA/firm_license.pdf']

        merge_pdfs(merge_lis, os.path.join(app_report_pat, thero[0] + '_LBP_Report_' + thero[11].strftime('%m%d%y') + '.pdf'))

        subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"])  # kill excel at end
