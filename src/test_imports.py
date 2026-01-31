import runpy

runpy.run_path('src/transport.py', run_name='src.transport')
runpy.run_path('src/gui.py', run_name='src.gui')
runpy.run_path('src/main.py', run_name='src.main')

print('Imported src modules OK')
