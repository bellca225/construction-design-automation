import tabula
from tabula import read_pdf
import json
from loguru import logger

#reads table from pdf file
data = {
    'twisted_cable' : {
        '4p' : 0
    },
    'connector_jack' : {
        'RS-232C(10Pin)' : 0,
        'Modular(RJ45-8PinPlug)': 0,
        'Modular(Outlet)':0,
        'TELCO(50Pin)' : 0,
        'Data Line':0,
    },
    'control_cable' : {
        '2C' : {
            '2.5mm' : 0
        }
    },
    'Communication_premises_power_cable' : {
        '25mm' : 0
    },
    'fr_cable' : 0,
}
def getData ():
    logger.info('Parsing pdf file..')
    df = read_pdf("./datas/standard_production.pdf",pages="all", stream=True) #address of pdf file
    for i in range(len(df)):
        dict = df[i].to_dict();
        # twisted_cable
        dumpp = json.dumps(dict)
        if '통 신' in dict:
            isFind = False
            for d in list(dict.values()) :
                for dd in list(d.values()) :
                    if type(dd) is str and dd.find('UTP') != -1:
                        isFind = True
            # print(str(isFind))
            # print(dict['통 신'])
            if isFind:
                data['twisted_cable']['4p'] = float(dict['통 신'][2]);
        # connector_jack
        elif 'RS-232C(10Pin)' in dumpp:
            t = dict['통신내선공'];
            data['connector_jack']['RS-232C(10Pin)'] = float(t[0])
            data['connector_jack']['Modular(RJ45-8PinPlug)'] = float(t[1])
            data['connector_jack']['Modular(Outlet)'] = float(t[2])
            data['connector_jack']['TELCO(50Pin)'] = float(t[3])
            data['connector_jack']['Data Line'] = float(t[4])
        # control_cable
        elif '1 C' in dumpp and '1 C' in dumpp:
            for d in list(dict.values()):
                if '2.5mm2' in json.dumps(d) :
                    data['control_cable']['2C']['2.5mm'] = float(d[2])
        elif '규격' in dict and 'P.V.C' in dumpp:
            _human = list(dict.values())[1]
            data['Communication_premises_power_cable']['25mm'] = float(_human[3])
        elif 'FR 케이블' in dumpp : 
            data['fr_cable'] = float(dict['통신케이블공'][0])
    # print(data)
    logger.info('Parsing pdf file done!')
    return data

# tabula.convert_into("abc.pdf", "output.csv", output_format="csv", pages='all')