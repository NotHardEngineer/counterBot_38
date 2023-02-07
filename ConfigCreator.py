import configparser

config = configparser.ConfigParser()

# Add the structure to the file we will create
config.add_section('program_settings')
config.set('program_settings', 'token', '')
config.set('program_settings', 'tablePath', r"")

config.add_section('table_info')
config.set('table_info', 'coldWaterSheetName', 'холодная')
config.set('table_info', 'hotWaterSheetName', 'горячая')
config.set('table_info', 'electricitySheetName', 'электроэнергия')
config.set('table_info', 'gasSheetName', 'газ')

config.add_section('counter_info')
config.set('counter_info', 'gasLength', '12')
config.set('counter_info', 'coldWaterLength', '13')
config.set('counter_info', 'hotWaterLength', '14')
config.set('counter_info', 'electricityLenght', '15')

# Write the new structure to the new file
with open("сonfigfile.ini", 'w', encoding='utf-8') as configfile:
    config.write(configfile)