# ReportProcessorHomepage.spec

# Required for encryption (leave as is unless you’re encrypting Python bytecode)
block_cipher = None

a = Analysis(
    ['Homepage.py'],   # Main script
    pathex=['.'],                     # Base path (project root)
    binaries=[],
    datas=[
        ('Pickupreportexe.py', '.'),       # Include module source files
        ('ReturnsReportexe.py', '.'),
        ('Cancellationexe.py', '.'),
        ('InputDIR', 'InputDIR'),          # Include your input/output folders
        ('Output', 'Output')
    ],
    hiddenimports=[
        'pandas',
        'pdfplumber',
        'openpyxl',
        'customtkinter',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.scrolledtext',
        'matplotlib',
        'Pickupreportexe',
        'ReturnsReportexe',
        'Cancellationexe'
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ReportProcessor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False  # Change to True if you want a terminal window (for debugging)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ReportProcessor'
)
