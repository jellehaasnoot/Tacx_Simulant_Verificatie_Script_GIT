# -*- mode: python -*-

block_cipher = None


a = Analysis(['SimulANT+ Log Analyzer.py'],
             pathex=['C:\\Users\\Jelle\\PycharmProjects\\Tacx_Simulant_Verificatie_Script'],
             binaries=[],
             datas=[],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          name='SimulANT+ Log Analyzer',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
