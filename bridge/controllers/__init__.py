"""Kontrolery COM - jeden na aplikacje Office."""

from bridge.controllers.base import BaseController
from bridge.controllers.excel import ExcelController
from bridge.controllers.powerpoint import PowerPointController

__all__ = ["BaseController", "ExcelController", "PowerPointController"]
