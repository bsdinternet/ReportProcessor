@echo off
echo Creating Report Processing Suite folder structure...

mkdir ReportProcessingSuite
cd ReportProcessingSuite

mkdir InputDIR
mkdir InputDIR\PickupReportfiles
mkdir InputDIR\Returnreportfiles
mkdir InputDIR\CancellationReport

mkdir Output
mkdir Output\Pivot_PNGs
mkdir Documentation
mkdir Dependencies

echo Folder structure created successfully!
echo.
echo Please copy your Python files to the ReportProcessingSuite directory:
echo - ReportProcessorHomepage.py
echo - Pickupreportexe.py
echo - Returnreportexe.py
echo - Cancellationexe.py
echo.
pause