# -*- mode: python -*-

block_cipher = None


a = Analysis(['weborder.py'],
             pathex=['/Users/Zach/Documents/projects/gateway-non-catalog-helper'],
             binaries=[('ChromeDriver', '.')],
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
          name='MacOSX-GatewayHelper',
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
               name='MacOSX-GatewayHelper')
