import requests
import logging
import json
from typing import Optional, Dict, Any, List

# Configure basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def search_google_books(query: str, author: Optional[str] = None, api_key: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """
    Search Google Books API for a book and return CSL-JSON compatible metadata.
    
    Args:
        query (str): The book title or general query.
        author (str, optional): The author's name to refine the search.
        api_key (str, optional): Google Books API key.
        
    Returns:
        dict: valid CSL-JSON item or None if no match found.
    """
    base_url = "https://www.googleapis.com/books/v1/volumes"
    params = {'q': query, 'maxResults': 1, 'printType': 'books'}
    
    if author:
        # Standardize author search
        params['q'] += f"+inauthor:{author}"
        
    if api_key:
        params['key'] = api_key
        
    try:
        logger.info(f"Querying Google Books: {params['q']}")
        resp = requests.get(base_url, params=params, timeout=10)
        
        if resp.status_code == 200:
            data = resp.json()
            if 'items' in data and len(data['items']) > 0:
                book = data['items'][0]['volumeInfo']
                
                # Extract and Normalize to CSL-JSON
                title = book.get('title', '')
                subtitle = book.get('subtitle', '')
                if subtitle:
                    title += f": {subtitle}"
                    
                # Handle Authors
                authors_list = []
                for a in book.get('authors', []):
                    # Google gives "FirstName LastName" usually
                    parts = a.split()
                    if len(parts) > 1:
                        authors_list.append({'given': " ".join(parts[:-1]), 'family': parts[-1]})
                    else:
                        authors_list.append({'literal': a})
                
                # Handle Date
                pub_date = book.get('publishedDate', '')
                date_parts = []
                if pub_date:
                    # YYYY-MM-DD or YYYY
                    year_match = pub_date[:4]
                    if year_match.isdigit():
                        date_parts = [[int(year_match)]]
                
                # Construct CSL Item
                csl_item = {
                    'title': [title],
                    'author': authors_list,
                    'published-print': {'date-parts': date_parts} if date_parts else {},
                    'publisher': book.get('publisher', ''),
                    'type': 'book',
                    'URL': book.get('infoLink', ''),
                    # Add standard 'DOI' field as empty since GBooks often doesn't have it, 
                    # but it helps downstream logic not crash
                    'DOI': ''
                }
                
                # Check for Industry Identifiers (ISBN)
                for ident in book.get('industryIdentifiers', []):
                    if ident['type'] == 'ISBN_13':
                        csl_item['ISBN'] = ident['identifier']
                        
                logger.info(f"Book found: {title}")
                return csl_item
            else:
                logger.info("No results found in Google Books.")
                return None
        else:
            logger.warning(f"Google Books API Error: {resp.status_code} - {resp.text}")
            return None
            
    except Exception as e:
        logger.error(f"Google Books request failed: {e}")
        return None

if __name__ == "__main__":
    # Test cases
    test_queries = [
        ("Marketing research: Methodological foundations", "Churchill"),
        ("Services marketing: Concepts, strategies, & cases", "Hoffman")
    ]
    
    print("--- Testing Book Validation ---")
    for q, a in test_queries:
        print(f"\nSearching for: '{q}' by '{a}'")
        result = search_google_books(q, a)
        if result:
            print("Match Found:")
            print(json.dumps(result, indent=2))
        else:
            print("No match found.")
