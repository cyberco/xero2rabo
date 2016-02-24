"""
Tool to convert Xero's exported bank statement (CSV) to a format Rabobank can
process SEPA Credit Transfer.

SEPA Credit Transfer format (also called pain.001.001.03):
https://www.rabobank.nl/images/formaatbeschrijving_xml_sepa_ct_v1_6_29541337.pdf

Xero's output format (CSV):
300.00,NL23RABO0,A. de Vries,Wages,Wages1 Feb - 28 Feb

An online tool for generating working SEPA documents can be found here:
http://www.mobilefish.com/services/sepa_xml/sepa_xml.php
"""
import csv
import xml.etree.ElementTree as ET
from argparse import ArgumentParser
from datetime import datetime
import random
import string
import copy

now = datetime.now()


def createMsgId(prefix):
    """Create a message ID conforming to the RABO spec."""
    id = prefix[:5] if len(prefix) > 4 else prefix.zfill(5)
    id += ('-{}' + '{:02d}' * 5 + '-').format(now.year, now.month, now.day, now.hour, now.minute,
            now.second)
    # Fill up to 35 chars
    id += ''.join(random.choice(string.ascii_lowercase + string.digits) for _ in range(7))
    return id


def get_credit_transactions(filename):
    with open(filename, 'rt') as input_file:
        reader = csv.reader(input_file)
        for row in reader:
            yield {'amount': row[0], 'iban': row[1], 'creditor_name': row[2], 'descr': row[4]}


def process_xml(args):
    """Process the SEPA template XML and fill with data."""
    tree = ET.parse('sepa-template.xml')
    # Register the namespaces so it will not clutter up all element names when printing. Makes not
    # difference for search.
    namespaces = {'': 'urn:iso:std:iso:20022:tech:xsd:pain.001.001.03',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}
    for k, v in namespaces.items():
        ET.register_namespace(k, v)
    # Main namespace needed for search
    namespace = namespaces['']
    root = tree.getroot()
    # Get all transactions
    credit_transactions = list(get_credit_transactions(args.input_file))
    # Set MsgId
    msgId = root.find('.//{%s}MsgId' % namespace)
    id = createMsgId(args.id_prefix)
    msgId.text = id
    # Set CreDtTm
    creDtTm = root.find('.//{%s}CreDtTm' % namespace)
    creDtTm.text = now.isoformat().split('.')[0]
    # Number of trx (NbOfTxs)
    NbOfTxs = root.find('.//{%s}NbOfTxs' % namespace)
    NbOfTxs.text = str(len(credit_transactions))
    # Initiating party
    initgPty = root.find('.//{{{0}}}InitgPty/{{{0}}}Nm'.format(namespace))
    initgPty.text = args.initiating_party
    # PaymentInformationIdentification
    pmtInfId = root.find('.//{%s}PmtInfId' % namespace)
    id += '-1'
    pmtInfId.text = id
    # Reguested execution date
    # Use today by default
    reqdExctnDt = root.find('.//{%s}ReqdExctnDt' % namespace)
    reqdExctnDt.text = now.date().isoformat()
    # Debtor name
    dbtr = root.find('.//{{{0}}}Dbtr/{{{0}}}Nm'.format(namespace))
    dbtr.text = args.debtor_name
    # Debtor account
    ibanDbtr = root.find('.//{{{0}}}DbtrAcct/{{{0}}}Id/{{{0}}}IBAN'.format(namespace))
    ibanDbtr.text = args.debtor_account
    # Debtor BIC
    dbtr_bic = root.find('.//{{{0}}}DbtrAgt/{{{0}}}FinInstnId/{{{0}}}BIC'.format(namespace))
    dbtr_bic.text = args.debtor_bic
    # Add credit transactions
    cdtTrfTxInf = root.find('.//{%s}CdtTrfTxInf' % namespace)
    parent = root.find('.//{%s}PmtInf' % namespace)
    parent.remove(cdtTrfTxInf)
    for index, tx in enumerate(credit_transactions):
        elem = copy.deepcopy(cdtTrfTxInf)
        # EndToEndId
        endToEndIdText = id + '-' + str(index).zfill(4)
        endToEndIdElem = elem.find('.//{%s}EndToEndId' % namespace)
        endToEndIdElem.text = endToEndIdText
        # InstdAmt
        instdAmt = elem.find('.//{%s}InstdAmt' % namespace)
        instdAmt.text = tx['amount']
        # Creditor name
        cdtr = elem.find('.//{{{0}}}Cdtr/{{{0}}}Nm'.format(namespace))
        cdtr.text = tx['creditor_name']
        # Creditor IBAN
        cdtr_iban = elem.find('.//{{{0}}}CdtrAcct/{{{0}}}Id/{{{0}}}IBAN'.format(namespace))
        cdtr_iban.text = tx['iban']
        # Ustrd
        ustrd = elem.find('.//{%s}Ustrd' % namespace)
        ustrd.text = tx['descr']
        parent.append(elem)
    return tree


if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument("input_file", help="Input file (CSV)")
    parser.add_argument("output_file", help="Output file (xml)")
    parser.add_argument('-id', '--id_prefix', default='', help='5-char prefix for the message ID')
    parser.add_argument('-ip', '--initiating_party', default='', help='Initiating party')
    parser.add_argument('-dn', '--debtor_name', default='', help='Debtor name')
    parser.add_argument('-da', '--debtor_account', default='', help='Debtor account')
    parser.add_argument('-db', '--debtor_bic', default='', help='Debtor BIC')
    args = parser.parse_args()
    tree = process_xml(args)
    # html output ensures emtpy elements don't use shorthand notation (<el />), which would break
    tree.write(args.output_file, method='html')
