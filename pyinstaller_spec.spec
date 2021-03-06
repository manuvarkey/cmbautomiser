# -*- mode: python -*-

block_cipher = None

added_files = [
         ( 'cmbautomiser/interface/*.glade', 'interface' ),
         ( 'cmbautomiser/templates/*.py', 'templates' ),
         ( 'cmbautomiser/latex/*.tex', 'latex' ),
         ( 'cmbautomiser/ods_templates/*.xlsx', 'ods_templates' ),
         ( 'cmbautomiser/documentation/*.pdf', 'documentation' ),
         ( 'miketex/', 'miketex' )
		 ]

a = Analysis(['cmbautomiser/main.py'],
             pathex=['C:\\Users\\User\\Desktop\\cmbautomiser\\cmbautomiser'],
             binaries=None,
             datas= added_files,
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='main',
          debug=False,
          strip=False,
          upx=False,
          console=False,
          icon = 'cmbautomiser.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               name='__init__')
