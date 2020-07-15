# ______________________________________________________________________________________________________________________
# ______________________________________________________________________________________________________________________
#
#
# BATCH MAILMERGE DOCX TEMPLATES WITH PYTHON
#
# ______________________________________________________________________________________________________________________
# ______________________________________________________________________________________________________________________
#
# File name: SampleMailMerge.py
# Author: Josh Edwards
# Email: Josh@JoshEdwardsGuitar.com
# Date created: 09/19/2018
# Last modified: 06/15/2020
# Python Version: 3.4
# Purpose: To save time and reduce work load from repetitive tasks related to document drafting. 
# Description: Program performs mailmerge on batch of documents by passing user input field declarations and automated derivatives into set of docx templates.
# Note: This script has been stripped down from it's original version to protect certain company data and procedures. 
# Use-Case: Used to batch draft legal documents associated with various types of real estate projects.
# ______________________________________________________________________________________________________________________


# IMPORTS
from __future__ import print_function
from mailmerge import MailMerge  # The package that makes this possible.
from datetime import date


# VARIABLES
# Define all variables, which are the "Merge Field" keys and values.
Declarant = 'John Doe and Jane Deere'
DeclarantType = 'individually'
DeclarantAddress = '3813 Avenue H, Austin, Texas 78751'
DeclarantEntity = ''
By = ''
Member1 = 'John Doe'
Mem1Title = 'MEM1TITLE'
Signature1 = 'John Doe'
Acknowledgement1 = 'Darren Boerner.'
Member2 = 'Jane Deere'
Mem2Title = 'MEM2TITLE'
Signature2 = 'Jane Deere'
Acknowledgement2 = 'Jennifer Boerner.'
S = 'S'
SignatureFollowing = '[Signatures appear on following pages.]'
President = 'John Doe'
Secretary = 'Jane Deere'
Treasurer = 'John Doe'
EffectiveDate = 'July 15, 2020'
HOAName = '1405 Main' # 1405 Main Condominiums
HOAAddress = '1405 Main Street, Austin, Texas 78704'
LegalDescription = 'Lot 11, Block J, FOREST OAKS SECTION 3, according to the map or plat thereof, recorded in Volume 11, Page 1111, Plat Records, Travis County, Texas.'
NumberOfUnits = 'two (2)'
BothOrAll = 'both'
AttachedOrDetached = 'detached'
AnAttachedOrADetached = 'a detached'
CondoType = 'The Condominium consists of two (2), detached Units in two (2) buildings. There shall be a maximum of two (2) Units.'
Unit1 = 'Unit 1'
Unit2 = 'Unit 2'
Unit1Sqft = '823'
Unit2Sqft = '1,500'
Unit1SqftExclusions = '(not including ______________________)'
Unit2SqftExclusions = '(not including ______________________)'
Unit1SqftSource = 'MLS listing'
Unit2SqftSource = 'architectural plans'
Fee = '$4,500.00'

# LOGIC
# Perform any additional logic, if desired, to make new or modified fields and field data.
# This may also be done in excel, Google Sheets, or inside the actual Microsoft Word merge fields.
'''
if DeclarantType == 'individually':
    DeclarantEntity = ''
    By = ''
else:
    DeclarantEntity = Declarant
    By = 'By:'

if DeclarantType == 'individually' and NumberOfMembers >= 2:
    S = 'S'
else:
    S = ''

if NumberOfUnits == 2:
    NumberOfUnits = 'two (2)'
    BothOrAll = 'both'
elif NumberOfUnits == 3:
    NumberOfUnits = 'three (3)'
    BothOrAll = 'all'
elif NumberOfUnits == 4:
    NumberOfUnits = 'four (4)'
    BothOrAll = 'all'

if AttachedOrDetached == 'attached':
    CondoType = 'The Condominium consists of ' + NumberOfUnits + ' Units within one (1) building, and there shall be a maximum of ' + NumberOfUnits + ' Units'
else:
    CondoType = 'The Condominium consists of two (2), detached Units in two (2) buildings. There shall be a maximum of two (2) Units. '
'''


# TEMPLATES
# Define which templates to be called. These are MS Word docx templates with Merge Fields already inserted. Note, this is also able to merge fields within the document footer.
# Template MUST live in same folder as script (unless naming the entire path).
def main():
    executeMailMerge('TEMPLATE temp1.docx')
    executeMailMerge('TEMPLATE temp2.docx')
    executeMailMerge('TEMPLATE temp3.docx')
    executeMailMerge('TEMPLATE temp4.docx')
    executeMailMerge('TEMPLATE temp5.docx')
    executeMailMerge('TEMPLATE temp6.docx')
    executeMailMerge('TEMPLATE temp7.docx')
    executeMailMerge('TEMPLATE temp8.docx')
    executeMailMerge('TEMPLATE temp9.docx')
    # add more templates here...


# MAILMERGE
def executeMailMerge(template):

    # Call the specific template.
    document = MailMerge(template)

    # Perform the merge on that specific template.
    document.merge(
        Declarant=Declarant,
        DeclarantType=DeclarantType,
        By=By,
        S=S,
        BothOrAll=BothOrAll,
        DeclarantEntity=DeclarantEntity,
        DeclarantAddress=DeclarantAddress,
        HOAName=HOAName,
        HOAAddress=HOAAddress,
        LegalDescription=LegalDescription,
        Unit1=Unit1,
        Unit2=Unit2,
        NumberOfUnits=NumberOfUnits,
        AttachedOrDetached=AttachedOrDetached,
        AnAttachedOrADetached=AnAttachedOrADetached,
        CondoType=CondoType,
        Unit1Sqft=Unit1Sqft,
        Unit2Sqft=Unit2Sqft,
        Unit1SqftExclusions=Unit1SqftExclusions,
        Unit2SqftExclusions=Unit2SqftExclusions,
        Unit1SqftSource=Unit1SqftSource,
        Unit2SqftSource=Unit2SqftSource,
        EffectiveDate=EffectiveDate,
        Member1=Member1,
        Mem1Title=Mem1Title,
        Signature1=Signature1,
        Member2=Member2,
        Mem2Title=Mem2Title,
        Signature2=Signature2,
        President=President,
        Secretary=Secretary,
        Treasurer=Treasurer,
        Acknowledgement1=Acknowledgement1,
        Acknowledgement2=Acknowledgement2,
        SignatureFollowing=SignatureFollowing,
        Fee=Fee)    


    # NAMING
    # Name the Document.
    document.write(HOAName + ' Condominiums - MERGED ' + template.replace('TEMPLATE ', ''))   
    #document.write('1205 West 36 Half Street' + ' Condominiums - MERGED ' + template.replace('TEMPLATE ', ''))      


# EXECUTE
# Call the main function, which executes mailmerge on each defined template.
main ()

















