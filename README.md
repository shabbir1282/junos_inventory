# junos_inventory
a script to convert junos show chassis hardware | display xml into excel sheet

this will generate a excel sheet in the same file with the name with same as input.xml including the date.

Dependencies:

xlsxwriter and lxml
  pip install xlsxwriter
  pip install lxml

USAGE: python get_inventory.py -f output.xml



