# Downloaded ORSA Response Files

This directory contains downloaded ORSA response files from document sourcing.

## Document Organization

Files are downloaded by the `ORSADocumentSourcer` based on the specified Berichtsjahr (reporting year):
- Default reporting year: 2026
- Configurable via `--berichtsjahr` command-line argument
- Each file is uniquely identified by its GeschaeftsNr (business number)

## Usage

Files in this directory are automatically:
1. Downloaded from the FINMA database
2. Filtered based on the specified Berichtsjahr
3. Processed through quality control checks
4. Tracked using hash-based caching to avoid reprocessing

## Note on Uniqueness

While the Berichtsjahr parameter helps filter documents for a specific reporting period, 
the GeschaeftsNr is unique for each institute and year combination. Over time, as multiple 
years of data accumulate, the GeschaeftsNr remains the primary unique identifier.
