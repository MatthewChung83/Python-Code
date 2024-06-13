import pdfplumber
import tempfile
import logging


class PDFReader:
    def __init__(self, file_path):
        self.file = pdfplumber.open(file_path)

    def extract(self):
        # 是5人 * 2組地號
        owners = []
        parcels = []
        parcels_tmp = []
        remarks = []
        country = ''
        town = ''
        for page in self.file.pages:
            for table in page.extract_tables():
                for row in table:
                    row = [r for r in row if r != None]

                    if not owners and '財產所有人：' in row[0]:
                        #if '、' in row[0]:
                        #    owners = row[0].split('財產所有人：')[-1].split('、')
                        #elif '，' in row[0]:
                        #    owners = row[0].split('財產所有人：')[-1].split('，')
                        #elif ',' in row[0]:
                        #    owners = row[0].split('財產所有人：')[-1].split(',')
                        #elif ' ' in row[0]:
                        #    owners = row[0].split('財產所有人：')[-1].split(' ')
                        #elif '兼' in row[0]:
                        #    owners = row[0].split('財產所有人：')[-1].split('兼')
                        #elif '即' in row[0]:
                        #    owners = row[0].split('財產所有人：')[-1].split('即')
                        #elif '之' in row[0]:
                        #    owners = row[0].split('財產所有人：')[-1].split('之')
                        
                        owners_temp = row[0].split('財產所有人：')[-1].replace('兼',',').replace('。',',').replace('/n',',').replace('、',',').replace('，',',').replace(' ',',').replace('：',',').replace('歿',',').replace('即',',').replace('之繼承人',',').replace('繼承人',',').replace('之遺產管理人',',').replace('遺產管理人',',').replace(':',',').replace('之限定繼承人',',').replace('限定繼承人',',').replace('原名',',').replace('權利範圍',',').replace('均為',',').replace('應有',',').replace('之有限責任繼承人',',').replace('有限責任繼承人',',').replace('之律師',',').replace('律師',',').replace('之清算管理人',',').replace('清算管理人',',').replace("（",',').replace("）",',').replace("(",',').replace(")",',').replace('之遺產管理人',',').replace('遺產管理人',',').split(',')
                        owners = list(filter(None, owners_temp))
                    if row[0].isdigit()and len(list(row))>2:
                        if '----------' not in row[2] and '、' not in row[2]:
                            # print(row)
                            # 坪頂段 二小段 248地號
                            # ['1', '新北市', '淡水區', '坪頂', '', '248', '1.12', '1分之1', '70,000元']
                            country = row[1].replace('臺', '台')
                            parcel_last = [k+v for k, v in zip(row[3:6],['段', '小段','地號']) if k]
                            par = {
                                'parcel': ''.join([country]+row[2:3]+parcel_last),
                                'area': ' x '.join(row[6:8]).replace('\n', '').strip(),
                            }
                            if par not in parcels:
                                parcels.append(par)
                        else:
                            # ['1', '1111', '新北市淡水區坪\n頂段249、250地\n號\n--------------\n新北市淡水區八\n勢路一段123巷1\n號', '鋼筋混\n凝土造5\n層', '一層: 133.30\n二層: 133.30\n三層: 130.74\n四層: 117.69\n五層: 90.11\n屋頂突出物: 31.57\n合計: 636.71', '陽台47.02', '2分之1', '7,810,000元']
                            remarks.append(row[1].strip().replace('\n', ''))

                            parcel = row[2].replace('\n', '').split('--------------')[0]
                            if '、' in parcel:
                                # 金門縣金湖鎮太 湖劃測段355、 356、362地號
                                pars = parcel.split('、')
                                par_tmp = pars[0]+'地號'
                                if par_tmp not in parcels_tmp:
                                    parcels_tmp.append(
                                        {'parcel': par_tmp}
                                    )

                                main_addr = pars[0].split('測段')[0] + '測段'
                                for par in pars[1:-1]:
                                    par_tmp = main_addr+par+'地號'
                                    if par_tmp not in parcels_tmp:
                                        parcels_tmp.append(
                                            {'parcel': par_tmp}
                                        )

                                par_tmp = main_addr+pars[-1]
                                if par_tmp not in parcels_tmp:
                                    parcels_tmp.append(
                                        {'parcel': par_tmp}
                                    )

                            else:
                                if parcel not in parcels_tmp:
                                    parcels_tmp.append({
                                        'parcel': parcel,
                                    })

        if not len(parcels):
            parcels = parcels_tmp

        output = []
        for owner in owners:
            for parcel in parcels:
                output.append({
                    'remark': ';'.join(remarks),
                    'owner': owner,
                    **parcel,
                })
        return output
