# -*- mode: python -*-

block_cipher = None


a = Analysis(['gatewayhelper.py'],
             pathex=['/Users/Zach/Documents/projects/gateway-non-catalog-helper'],
             binaries=[('ChromeDriver', '.')],
             datas=[('ui_gatewayhelper.py', '.'), ('settings.json', '.')],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='MacOSX-GatewayHelper',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='icon.icns')
app = BUNDLE(exe,
             name='MacOSX-GatewayHelper.app',
             icon='icon.icns',
             bundle_identifier=None)
