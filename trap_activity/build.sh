#!/bin/bash
pyinstaller --add-data 'seasons.json:.' trap_activity_gui.py
