import pandas as pd
import numpy as np
import codecs
import json
import time
import hashlib
import re
from neo4j import GraphDatabase


def start_neo4j_api():
    with open('./Input/config.json') as json_file:
        config = json.load(json_file)
    driver = GraphDatabase.driver(config["Server"], auth=(config["Server_Login"], config["Server_PW"]))
    return driver

# ################################################################################################################
# ###############    Add common attributes     ###################################################################
# ################################################################################################################

common_attributes = []
common_attributes.append('Details')
common_attributes.append('Description')
common_attributes.append('DeepLink')
common_attributes.append('Key')
common_attributes.append('Owner')
common_attributes.append('ModifyDate')


driver = start_neo4j_api()

with driver.session() as session:
    for head in common_attributes:
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (n) WHERE NOT EXISTS(n.' + head + ')')
        qry_cmd_Lst.append('SET n.' + head + ' = \"\"')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)
        records = session.run(qry)
    session.close()


# ################################################################################################################
# ###############    Set Key where not set     ###################################################################
# ################################################################################################################
lst = []
with driver.session() as session:
    qry_cmd_Lst = []
    qry_cmd_Lst.append('MATCH (x) WHERE x.Key = ""')
    qry_cmd_Lst.append('SET x.Key = apoc.create.uuid()')
    qry = ' '.join(qry_cmd_Lst)
    print(qry)
    records = session.run(qry)


# ################################################################################################################
# ###############    Set Owner and Sync Property where not set     ###############################################
# ################################################################################################################
lst = []
with driver.session() as session:
    qry_cmd_Lst = []
    qry_cmd_Lst.append('MATCH (x) WHERE x.Owner = ""')
    qry_cmd_Lst.append('SET x.Owner = \"System\"')
    qry = ' '.join(qry_cmd_Lst)
    print(qry)
    records = session.run(qry)
session.close()

with driver.session() as session:
    qry_cmd_Lst = []
    qry_cmd_Lst.append('MATCH (n:Folder) WHERE NOT EXISTS(n.Sync)')
    qry_cmd_Lst.append('SET n.Sync = \"False\"')
    qry = ' '.join(qry_cmd_Lst)
    print(qry)
    records = session.run(qry)
    qry_cmd_Lst = []
    qry_cmd_Lst.append('MATCH (m:File) WHERE NOT EXISTS(m.Sync)')
    qry_cmd_Lst.append('SET m.Sync = \"False\"')
    qry = ' '.join(qry_cmd_Lst)
    print(qry)
    records = session.run(qry)
session.close()


# ################################################################################################################
# ###############    Löscht Knoten ohne HookNode   ###############################################################
# ################################################################################################################

with driver.session() as session:
    Label_lst = []
    qry1 = ('MATCH (n) RETURN distinct labels(n) ')

    records = session.run(qry1)

    for record in records:
        print(record['labels(n)'])
        for tmp in record['labels(n)']:
            Label_lst.append(tmp)
    for Lable_tmp in Label_lst:
        print(Lable_tmp)
    Label_lst = sorted(Label_lst, key=str.casefold)

    for label in Label_lst:
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (x:HookNode)')
        qry_cmd_Lst.append('MATCH (x)-[*0..15]->(w:' + label + ')')
        qry_cmd_Lst.append('WITH COLLECT(w) AS wanted')
        qry_cmd_Lst.append('MATCH (y:' + label + ')')
        qry_cmd_Lst.append('WHERE NOT y IN wanted')
        qry_cmd_Lst.append('DETACH DELETE y')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)
        session.run(qry)
session.close()



driver.close()