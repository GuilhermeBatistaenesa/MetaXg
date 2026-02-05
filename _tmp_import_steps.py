import time
steps = [
    'pyodbc',
    'custom_logger',
    'output_manager',
    'outcomes',
    'rpa_metax',
    'sharepoint',
    'config',
    'notification',
    'reporting',
]
with open('import_steps.log','w') as f:
    for name in steps:
        f.write(f'start {name}\n')
        f.flush()
        __import__(name)
        f.write(f'ok {name}\n')
        f.flush()
