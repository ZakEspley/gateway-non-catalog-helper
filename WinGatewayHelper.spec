# -*- mode: python -*-

block_cipher = None


a = Analysis(['gatewayhelper.py'],
             pathex=['C:\\Users\\Instlab\\Documents\\gateway-non-catalog-helper'],
             binaries=[('chromedriver.exe', '.')],
             datas=[('ui_gatewayhelper.py', '.'), ('settings.json', '.'), ('icon.ico', '.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=['tk'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='GatewayHelper',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='icon.ico')
