# -*- coding: utf-8 -*-

import xmlrpclib
import csv
from pprint import pprint
import argparse
import sys
import logging
from logging import debug, info, warning
import os
from pprint import pprint

##GLOBAL VARIABLES
default = {'name': '','default_code': '','ean13':'','standard_price': 0,'list_price':0,'is_product_variant': True}
#templates : {template: [index de productes amb aquest template]}
templates = {}
#eans: {ean: {diccionari del row}}
eans = {}
#h : {cabecera: index col}
h = {}
#r : {index: {diccionari de row}}
r = {}
#default_template {template: {diccionari de valors que hi ha a la primera linia}}
default_template = {}
#att {attribut: ""} Llista d'attributs, posteriorment els ids a que equivalen.
att = {}
#t2att {template: [llista de attributs pel template]}
t2att = {}
#template_id {template: id de la databse}
template_id = {}
#template_id {variant: id de la databse}
product_id = {}
#guarda el id del attribute_line
attribute_line_id = ""
rows = []
tmp_tmpl = ""


def parse_args():
    parser = argparse.ArgumentParser(description="Import csv file to odoo server.",fromfile_prefix_chars='=')
    parser.add_argument("file", help='Name of csv file to open. (, and " formated)')
    parser.add_argument("-url", help="URL of the site.",required=True)
    parser.add_argument("-db", help="Database to connect.",required=True)
    parser.add_argument("-user", help="User.",required=True)
    parser.add_argument("-password", help="Password.",required=True)
    return parser.parse_args()


def processar_fitxer():
    for idx, row in enumerate(rows):
        if idx == 0:
            #first row - Mapa de camps
            for idx2, camp in enumerate(row):
                h[camp] = idx2
        else:
            # mirar si hi ha template
            if row[h['template']].strip():
                if row[h['template']] in templates:
                    # Si te template indicat afegeix
                    templates[row[h['template']]].append(idx)
                else:
                    # en cas contrari inicia la llista i guarda el tmp per futura referencia.
                    templates[row[h['template']]] = [idx]
                    tmp_tmpl = row[h['template']]
                    t2att[row[h['template']]] = []
                    default_template[row[h['template']]] = { k: v for (k,v) in zip(h.keys(),[ row[h[x]] for x in h.keys()]) }
            elif row[h['template']].strip() == "" and tmp_tmpl:
                # Si no hi ha template, te el valor guardat a tmp_tmpl
                templates[tmp_tmpl].append(idx)
            else:
                print "Error en linia", idx

            # crear el diccionari de rows.
            r[idx] = { k: v for (k,v) in zip(h.keys(),[ row[h[x]] for x in h.keys()]) }
            r[idx]['template'] = tmp_tmpl
            eans[row[h['ean13']]] = r[idx]
            if 'attribute_value' in h:
                att[r[idx]['attribute_value']] = r[idx]['attribute_id']
                if r[idx]['attribute_value'] not in t2att[tmp_tmpl]:
                    t2att[tmp_tmpl].append(r[idx]['attribute_value'])
    return


def crear_template(d):
    print "amb nom {0}.".format(d['name'])
    value = models.execute_kw(db,uid,password,'product.template','search',[[['default_code','=',d['template']]]])
    if value:
        print "Template {0} ja existent!".format(d['name'])
        return value[0]
    else:
        return models.execute_kw(db, uid, password, 'product.template', 'create', [d])

def crear_product(idx):
    print "Creant variant producte {0} {1}.".format(r[idx]['name'],r[idx]['attribute_value'])
    value = models.execute_kw(db,uid,password,'product.product','search',[[['default_code','=',r[idx]['default_code']]]])

    if value:
        print "Variant {0} ja existent!".format(r[idx]['name'])
        return value[0]
    else:
        d = default.copy()
        d.update(r[idx])
        for x in ['categ_id','list_price','standard_price']:
            #possar defaults en cas necesari
            if not d[x]: d[x] = default_template[d['template']][x]

        d['attribute_line_ids'] = [(6,0,[attribute_line_id])]
        d['attribute_value_ids'] = [(6,0,att[r[idx]['attribute_value']])]
        d['product_tmpl_id'] = template_id[r[idx]['template']]
        for x in ['attribute_id','template','attribute_value','name']:
            #Treure -> Template / attribute_value / attribute_id
            d.pop(x,None)
        # pprint (d)
        return models.execute_kw(db, uid, password, 'product.product', 'create', [d])

def crear_attributo(attrib):
    d = {
        "attribute_id": att[attrib],
        "display_name": attrib,
        "name": attrib
        }
    # pprint (d)
    value = models.execute_kw(db,uid,password,'product.attribute.value','search',[[['name','=',d['name']]]])
    if value:
        logging.info("Attributo '{0}' ya creado".format(d['name']))
    else:
        value = models.execute_kw(db, uid, password, 'product.attribute.value', 'create', [d])
    return value

def crear_atts_template(t):
    a_ids = []
    for v in t2att[t]:
        if v in att:
            a_ids.append(att[v][0])
        else:
            print "Error, no {v} en atts: {1}".format(v,att)
    d = {
        'attribute_id' : int(r[1]['attribute_id']),
        'product_tmpl_id' : template_id[t],
        'value_ids' :  [(6,0,a_ids)]
        }
    search = [['product_tmpl_id','=',d['product_tmpl_id']]] #,['value_ids','=',sorted(a_ids)]
    # print search
    value = models.execute_kw(db,uid,password,'product.attribute.line','search',[search])
    if value:
        #ja existeix
        print "Variants a template ja creades!", t, 'amb id', value
        # pprint  (models.execute_kw(db,uid,password,'product.attribute.line','read',[value],{'fields':[]}))
        return value[0]
    else:

        # pprint(d)
        value = models.execute_kw(db, uid, password, 'product.attribute.line', 'create', [d])
        print "Creant variants a template, id: ", value
        # pprint  (models.execute_kw(db,uid,password,'product.attribute.line','read',[value],{'fields':[]}))

        return value

def open_csv(filename):
    # Opens the sheet xls and gets the info needed to work
    if not os.path.isfile(filename):
        print "ERROR: File '{0}' not found.".format(filename)
        exit()
    return open(filename,'r')



if __name__ == "__main__":
    args = parse_args()
    logname = sys.argv[0]+'.log'
    with open(logname, 'w'):
        pass
    logging.basicConfig(format='%(levelname)s:%(message)s',filename=logname,level=logging.DEBUG)

    #Crear tota l'informació
    csv_file = open_csv(args.file)
    fitxer = csv.reader(csv_file,delimiter=',', quotechar='"')
    for x in fitxer:
        rows.append(x)
    processar_fitxer()

    #Conectar per XMLRPC al servidor
    common = xmlrpclib.ServerProxy(str(args.url)+'/xmlrpc/2/common')
    texte = common.version()
    uid = common.authenticate(args.db, args.user, args.password, {})
    models = xmlrpclib.ServerProxy('{}/xmlrpc/2/object'.format(args.url))
    db = args.db; password = args.password;

    #Resum de dades y confirmació.
    existeix = 0
    print "Hi ha {0} plantilles per introduir".format(len(templates))

    print "Hi ha {0} articles per introduir".format(len(r))
    llista = models.execute_kw(db,uid,password,'product.product','search_read',[],{'fields': ['default_code']})
    tots = [ (str(x['default_code'])) for x in llista]
    for x in r:
        if r[x]['default_code'] in tots: existeix += 1
    print "Dels que ja existeixen {0} i es donaran d'alta {1}.".format(existeix,len(r)-existeix)

    existeix = 0
    print "Hi ha {0} attributs diferents.".format(len(att))
    llista = models.execute_kw(db,uid,password,'product.attribute.value','search_read',[],{'fields': ['name','attribute_id']})
    tots = [ (str(x['attribute_id'][0]) , x['name']) for x in llista]
    for attributo in att:
        if (att[attributo],attributo) in tots: existeix += 1
    print "Dels que ja existeixen {0} i es donaran d'alta {1}.".format(existeix,len(att)-existeix)
    result = raw_input("Introdueixi 'SI' per continuar: ")
    if str(result).strip().upper() != "SI":
        print "Finalitzant programa, cap dada introduida..."
        exit()

    #Crear els Attributs
    for attributo in att:
        att[attributo] = crear_attributo(attributo)

    #Crear els templates
    for template in templates:
        #primer creem el template
        print "Creant template", template,
        d = default.copy()
        d.update(default_template[template])
        # el template no te ean13 ni default code
        d['ean13'], d['default_code'], d['is_product_variant'] = "", template, False
        template_id[template] = crear_template(d)
        # després hem de crear el product.attribute.line indicant els attributs d'aquest template
        attribute_line_id = crear_atts_template(template)
        # print "attribute_line_id", attribute_line_id
        # després hem de crear els product.product corresponents.
        for v in templates[template]:
            product_id[v] = crear_product(v)
        #Hem de borrar el variant que crea el template per defecte
        valor = models.execute_kw(db,uid,password,'product.product','search',[[['product_tmpl_id','=',template_id[template]]]])
        # print "Variant per defecte", valor
        # print "Variant per defecte primera", min(valor)
        # i borrar el mes baix.
        if valor: models.execute_kw(db,uid,password,'product.product','unlink',[[min(valor)]])

    import json
    debug("headers:")
    debug(json.dumps(h,indent=4,separators=(',', ': ')))
    debug("templates:")
    debug(json.dumps(templates,indent=4,separators=(',', ': ')))
    debug("t2att:")
    debug(json.dumps(t2att,indent=4,separators=(',', ': ')))
    debug("att:")
    debug(json.dumps(att,indent=4,separators=(',', ': ')))
    debug("template_id:")
    debug(json.dumps(template_id,indent=4,separators=(',', ': ')))
    debug("product_id:")
    debug(json.dumps(product_id,indent=4,separators=(',', ': ')))
