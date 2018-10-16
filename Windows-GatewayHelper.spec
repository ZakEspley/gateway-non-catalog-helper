# -*- mode: python -*-

block_cipher = None


a = Analysis(['weborder.py'],
             pathex=['C:\\Users\\Instlab\\Documents\\gateway-non-catalog-helper'],
             binaries=[('chromedriver.exe', '.')],
             datas=[('Carts', 'Carts'), ('Carts/NonCatalogTemplate.xlsx', 'Carts')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='Windows-GatewayHelper',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='Windows-GatewayHelper')
