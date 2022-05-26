cwd = os.path.abspath(os.path.dirname(__file__))
lpat = os.path.join('job_Folders', beholden[0] + ' - ' + beholden[5] + ' - ' + beholden[6], beholden[2] + '_LBP',
                    'app_Data')
lead_pat = os.path.abspath(os.path.join(cwd, lpat))
page_len = round(
    4 + len(os.listdir(os.path.join(lead_pat, 'elevations')))) / 6  # number of pages needed to accomadate photos