# BIS Document Intelligence & Governance Analysis System

**AI/ML-Powered RAG Pipeline for Banking Regulatory Document Analysis**

An advanced Retrieval-Augmented Generation (RAG) system that ingests, indexes, and analyzes large corpora of regulatory documents from the Bank for International Settlements (BIS). The system performs automated governance analysis, regulatory anomaly detection, and compliance gap identification across thousands of pages of financial policy documents.

## Capabilities

| Analysis Module | Description |
|----------------|-------------|
| Governance & Compliance Analysis | Identifies governance structures, oversight mechanisms, and compliance patterns |
| Regulatory Anomaly Detection | Flags inconsistencies, contradictions, and deviations across document corpus |
| Best Practices Identification | Extracts and ranks regulatory best practices from the full document set |
| Comprehensive Intelligence Report | Generates structured DOCX reports with findings and recommendations |
| Legal & Procedural Gap Analysis | Identifies gaps between stated procedures and regulatory requirements |
| Market Irregularity Assessment | Detects references to market irregularities and systemic risk indicators |

## Data Scale

- **156 BIS papers** analyzed (March 2001 – April 2025)
- **22,296 pages** fully indexed, processed, and vectorized
- Documents stored in SQLite with full-text search and TF-IDF vectorization

## Technical Stack

- **NLP Pipeline**: spaCy, NLTK (tokenization, lemmatization, stopword removal), TextBlob (sentiment analysis), textstat (readability scoring)
- **ML/Vectorization**: scikit-learn (TF-IDF, K-Means clustering, Latent Dirichlet Allocation for topic modeling)
- **Document Processing**: PyMuPDF (fitz), PyPDF2 for robust multi-format PDF extraction
- **Report Generation**: python-docx for automated DOCX report creation
- **GUI**: Tkinter-based desktop interface for interactive analysis
- **Storage**: SQLite for document indexing and metadata management

## Architecture

```
bisv3.py           # Latest version — full RAG pipeline with GUI and all 6 analysis modules
bisv2.py           # Prior iteration with core analysis functionality  
bisv1.py           # Initial implementation
BIS10/             # Sample subset of BIS papers (PDF) and generated reports (DOCX)
```

## Dependencies

```
PyPDF2
PyMuPDF
pandas
numpy
scikit-learn
nltk
spacy
textstat
textblob
python-docx
```

## Usage

```bash
pip install PyPDF2 PyMuPDF pandas numpy scikit-learn nltk spacy textstat textblob python-docx
python -m spacy download en_core_web_sm
python bisv3.py
```

The GUI will launch. Load PDF documents via the file browser, wait for indexing to complete, then select any of the six analysis modules to generate findings.

## Methodology

1. **Document Ingestion**: Batch PDF extraction with fallback between PyMuPDF and PyPDF2 for maximum compatibility
2. **Text Processing**: Sentence/word tokenization, lemmatization, stopword removal, named entity recognition via spaCy
3. **Vectorization**: TF-IDF vectorization of document segments for similarity search and retrieval
4. **Topic Modeling**: Latent Dirichlet Allocation (LDA) for unsupervised topic discovery across the corpus
5. **Clustering**: K-Means clustering of document segments by semantic similarity
6. **Analysis & Reporting**: Module-specific analytical pipelines generating structured DOCX reports

## Author

**Dr. Mosab Hawarey**
PhD, Geodetic & Photogrammetric Engineering | MSc, Geomatics (Purdue) | MBA (Wales)

- GitHub: [github.com/mhawarey](https://github.com/mhawarey)
- ORCID: [0000-0001-7846-951X](https://orcid.org/0000-0001-7846-951X)

## License

MIT License
