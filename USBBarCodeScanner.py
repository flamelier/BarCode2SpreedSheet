# BSD 3-Clause License

# Copyright (c) 2021, Flamelier
# All rights reserved.

# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:

# 1. Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.

# 2. Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.

# 3. Neither the name of the copyright holder nor the names of its
#    contributors may be used to endorse or promote products derived from
#    this software without specific prior written permission.

# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import datetime

workBook = Workbook()

# print(datetime.datetime.now()) # Debug

dateTimeValue = str(datetime.datetime.now())

dateValue = dateTimeValue[:10]

# print(dateValue) # Debug
userName = input('\n\nMAKE SURE TO USE A UNQUIE NAME OTHERWISE THE OLD FILE WILL BE OVER WRITTEN.\n\nName or Unquie Value: ')

fileName = 'Scanned=' + dateValue + '_by=' + str(userName) + '.xlsx'

workSheetOne = workBook.active
workSheetOne.title = "Scanned Data"

rowNumber = 1
print('\nFile name will be: ' + fileName)
print('\nTable name will be: Scanned Data\n\nInputed data will start on row 1.\n')
print('Type "save" to save the data.\nType "quit" to save and quit the program loop.\n\nMAKE SURE TO USE A UNQUIE NAME OTHERWISE THE OLD FILE WILL BE OVER WRITTEN.\n\n')
while True:
    serialNumber = input('Scan a barcode:')
    editedSerialNumber = serialNumber.upper().strip()
    if editedSerialNumber == 'DONE':
        workBook.save(filename = fileName)
    elif editedSerialNumber == 'SAVE':
        workBook.save(filename = fileName)
        print('Saving file.\n\n\n\n\n')
    elif editedSerialNumber == 'QUIT':
        workBook.save(filename = fileName)
        break
    else:
        # print('Entered value :' + editedSerialNumber) # Debug
        workSheetOne['A'+str(rowNumber)] = editedSerialNumber
        rowNumber +=1