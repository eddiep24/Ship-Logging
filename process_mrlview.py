import pandas as pd

def process_mrlview(mrlviewpath, mrldf):
    print("Opening mrlview file.")
    mrlviewdf = pd.read_excel(mrlviewpath)
    print("Opening mrl file.")
    # mrldf = pd.read_excel(mrlpath)
    n = 0
    pcount = 0
    matches = 0
    PSPindexes = []

    temphold = []

    for i in range(len(mrldf.index)):
        if mrldf.loc[i, 'MATCH'] == 'P' or mrldf.loc[i, 'MATCH'] == 'PS':
            print("Deleted", matches)
            matches += 1
            PSPindexes.append(i)
        print("Comparison", i, "out of", len(mrldf.index), "complete.")
    print(matches)
    mrldf = mrldf.drop(mrldf.index[PSPindexes])
    print(len(mrldf.index))
    for n in range(len(mrlviewdf.index)):
        # if mrlviewdf.loc[n, 'Match'] == 'PS' or mrlviewdf.loc[n, 'Match'] == 'P':
        mrlviewdf.loc[n, 'Activity ID'] = str(mrlviewdf.loc[n, 'Activity ID'])
        try:
            specsplit = mrlviewdf.loc[n, 'Activity ID'].split('.')
            if len(specsplit) == 2:
                for j in range(10):
                    if mrlviewdf.loc[n, 'SPEC'].startswith(str(j)):
                        try:
                            int(specsplit[1])
                            pval = 'P'
                        except:
                            if specsplit[1][-2:] == 'KE' or specsplit[1][-2:] == 'IM':
                                pval = None
                            else:
                                pval = 'PS'
            else:
                if int(specsplit[0]) == 1:
                    pval = None
                else:
                    pval = 'PS'
        except Exception:
            pval = 'PS'

        if pval == 'P':
            pcount += 1

        mrldf = mrldf.append({
                'MATCH' : pval,
                'PCT' : mrlviewdf.loc[n, 'Progress'], # C
                'Activity ID' : mrlviewdf.loc[n, 'Activity ID'], # D
                'Work Item' : mrlviewdf.loc[n, 'SPEC'], # E
                'TITLE' : mrlviewdf.loc[n, 'Activity Desc.'], # F
                'Ship Sup' : mrlviewdf.loc[n, 'SUPT\'D'], # Q
                'Dept/Shop #' : mrlviewdf.loc[n, 'resp'], # G
                'Dept/Shop #2' : mrlviewdf.loc[n, 'resp'], # G Number 2
                'Spec' : mrlviewdf.loc[n, 'Spec Para'], # W
                'Early/\nActual\nStart' : mrlviewdf.loc[n, 'Early Start'], # H
                'Early/\nActual\nStop' : mrlviewdf.loc[n, 'Early Finish'], # I
                'Late Start' : mrlviewdf.loc[n, 'Late Start'], # J
                'Late Stop' : mrlviewdf.loc[n, 'Late Finish'], # K
                'Duration' : mrlviewdf.loc[n, 'Dur'], # N
                'Total Float' : mrlviewdf.loc[n, 'Total Float'], # V
                'KE' : mrlviewdf.loc[n, 'Key Events'], # O
                'MS' : mrlviewdf.loc[n, 'IM / KE'], # P
                'Location' : mrlviewdf.loc[n, 'LOCATION'], # R
                'KE/MS SYSTEM' : mrlviewdf.loc[n, 'System'], # S
                'Key Trade (BAE)' : mrlviewdf.loc[n, 'SHOP'], # U
                'RR# \n(for PS task line)' : mrlviewdf.loc[n, 'Required Report Number'], # X
                'RR #' : mrlviewdf.loc[n, 'Required Report Number'], # X
                '009-67 ITP REQ\'D' : mrlviewdf.loc[n, 'WC TEST'], # Y
                }, ignore_index=True)

    print("Writing data to new file.")
    print(pcount, "Ps in the sheet")
    print(temphold)
    return mrldf


