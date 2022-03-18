import win32com.client as win32 
import pandas as pd 

def cleanCRNs(excel_df: pd.DataFrame) -> pd.DataFrame:
    crnList = []

    for index, row in excel_df.iterrows():
        crnList.append(((str)(row["사업자등록번호"])).strip())

    excel_df["사업자등록번호"] = crnList
    return excel_df

def matchCRN(path1, path2, outputDir):
    excel1 = cleanCRNs(pd.read_excel(path1))
    excel2 = cleanCRNs(pd.read_excel(path2))

    nameList = []
    addressList = []
    postcodeList = []

    for crn in excel1["사업자등록번호"]:
        targetRow = excel2.loc[excel2["사업자등록번호"] == crn]

        if targetRow.empty:
            nameList.append("일치하는 사업자 없음")
            addressList.append("일치하는 사업자 없음")
        else:
            nameList.append(targetRow.iloc[0]["대표자명"].strip())
            targetAddress = ((str) (targetRow.iloc[0]["대표자주소"])).strip()
            targetPostcode = ((str) (targetRow.iloc[0]["대표자우편번호"])).strip()

            if targetAddress == "nan" or targetPostcode == "null":
                addressList.append((str) (targetRow.iloc[0]["소재지"]).strip())
                postcodeList.append("확인필요")
            else:
                while (len(targetPostcode) < 5):
                    targetPostcode = "0" + targetPostcode

                addressList.append(targetAddress)
                postcodeList.append(targetPostcode)

    excel1.insert(len(excel1.columns), "대표자명", nameList, True)
    excel1.insert(len(excel1.columns), "대표자주소", addressList, True)
    excel1.insert(len(excel1.columns), "대표자우편번호", postcodeList, True)

    return excel1

def editHWP(excel, hwpFile):
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.Open(hwpFile)

    field_list = [i for i in hwp.GetFieldList().split("\x02")]

    # Select and Copy All Text
    hwp.Run('SelectAll')
    hwp.Run('Copy')

    # Move Cursor to End of Document
    hwp.MovePos(3)

    for i in range(len(excel) - 1):
        # Paste Previously Copied Text
        hwp.Run('Paste')

        # Move Cursor to End of Document
        hwp.MovePos(3)

    for page in range(len(excel)):
        for field in field_list:
            # Replace Fields with Data in Fields
            hwp.PutFieldText(f'{field}{{{{{page}}}}}', excel[field].iloc[page])

    hwp.Save()
    hwp.Quit()