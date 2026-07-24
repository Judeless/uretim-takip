@echo off
REM === Cofle Teyit Agent — PCOMM (RDP) oturumunda calisir ===
REM Server'da promanage Baslangic klasorune KISAYOL koy (shell:startup) —
REM logonda otomatik kalkar. Pencere acik kalmali (robot kouslari burada gorunur).
title Cofle Teyit Agent
cd /d %~dp0
python teyit_agent.py
pause
