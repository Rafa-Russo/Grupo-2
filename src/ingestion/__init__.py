"""Rotinas de ingestão de dados."""

from .crime import load_crime_data
from .listings import load_listings_data

__all__ = [ "load_listings_data"]