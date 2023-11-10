# -*- mode: python -*-

block_cipher = None

added_files = [
         ( 'cmbautomiser/interface/*.glade', 'cmbautomiser/interface' ),
         ( 'cmbautomiser/templates/*.py', 'cmbautomiser/templates' ),
         ( 'cmbautomiser/latex/*.tex', 'cmbautomiser/latex' ),
         ( 'cmbautomiser/ods_templates/*.xlsx', 'cmbautomiser/ods_templates' ),
         ( 'cmbautomiser/documentation/*.pdf', 'cmbautomiser/documentation' ),
         ( 'miketex/', 'cmbautomiser/miketex' )
		 ]

a = Analysis(['cmbautomiser_launcher.py'],
             pathex=['C:\\msys64\\home\\HP\\cmbautomiser\\cmbautomiser'],
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
