# -*- mode: python -*-

block_cipher = None


a = Analysis(['test_tkinter2.py'],
             pathex=['C:\\Users\\jinfeng\\eclipse-workspace\\xinxishuruxitong'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=True,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='test_tkinter2',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
