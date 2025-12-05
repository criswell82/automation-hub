"""
OneNote COM Automation Client.
Provides low-level COM interface to OneNote desktop application.
"""

import logging
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any
from datetime import datetime

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


class OneNoteCOMClient:
    """
    COM-based OneNote automation client.
    Uses win32com to interact with OneNote desktop application.
    """

    # OneNote XML namespace
    NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

    # Hierarchy scopes
    SCOPE_NOTEBOOKS = "{0ECB54FB-830B-4E61-880D-E45CB0E64B5D}"  # All notebooks
    SCOPE_SECTIONS = "{0FCE5E3B-4E61-468D-9E61-C62E5E5E5E5E}"  # All sections
    SCOPE_PAGES = "{0FCE5E3B-4E61-468D-9E61-C62E5E5E5E5F}"    # All pages

    def __init__(self):
        """Initialize COM client."""
        if not WIN32COM_AVAILABLE:
            raise RuntimeError("pywin32 is required for OneNote COM automation")

        self.logger = logging.getLogger(__name__)
        self.onenote = None
        self._connected = False

    def connect(self) -> bool:
        """
        Connect to OneNote application.

        Returns:
            bool: True if connection successful
        """
        try:
            self.onenote = win32com.client.Dispatch("OneNote.Application")
            self._connected = True
            self.logger.info("Connected to OneNote via COM")
            return True
        except Exception as e:
            self.logger.error(f"Failed to connect to OneNote: {e}")
            self._connected = False
            raise

    def is_connected(self) -> bool:
        """Check if connected to OneNote."""
        return self._connected

    def _ensure_connected(self):
        """Ensure OneNote connection exists."""
        if not self._connected:
            self.connect()

    # Hierarchy Navigation Methods

    def get_hierarchy(self, scope: str = None, xml_out: bool = False) -> Dict[str, Any]:
        """
        Get OneNote hierarchy (notebooks, sections, pages).

        Args:
            scope: Hierarchy scope (notebooks, sections, pages)
            xml_out: If True, return raw XML string instead of parsed dict

        Returns:
            dict or str: Hierarchy information
        """
        self._ensure_connected()

        # Default to notebooks scope
        if scope is None:
            scope = ""

        try:
            # Get XML hierarchy from OneNote
            xml_hierarchy = self.onenote.GetHierarchy("", 4)  # 4 = get all levels

            if xml_out:
                return xml_hierarchy

            # Parse XML to dict
            return self._parse_hierarchy_xml(xml_hierarchy)

        except Exception as e:
            self.logger.error(f"Failed to get hierarchy: {e}")
            raise

    def get_notebooks(self) -> List[Dict[str, str]]:
        """
        Get all notebooks.

        Returns:
            list: List of notebooks with ID, name, and path
        """
        hierarchy = self.get_hierarchy()
        return hierarchy.get('notebooks', [])

    def get_sections(self, notebook_id: str = None) -> List[Dict[str, str]]:
        """
        Get sections, optionally filtered by notebook.

        Args:
            notebook_id: Optional notebook ID to filter by

        Returns:
            list: List of sections
        """
        hierarchy = self.get_hierarchy()
        sections = []

        for notebook in hierarchy.get('notebooks', []):
            if notebook_id and notebook['id'] != notebook_id:
                continue

            for section in notebook.get('sections', []):
                section['notebook_id'] = notebook['id']
                section['notebook_name'] = notebook['name']
                sections.append(section)

        return sections

    def get_pages(self, section_id: str = None) -> List[Dict[str, str]]:
        """
        Get pages, optionally filtered by section.

        Args:
            section_id: Optional section ID to filter by

        Returns:
            list: List of pages
        """
        hierarchy = self.get_hierarchy()
        pages = []

        for notebook in hierarchy.get('notebooks', []):
            for section in notebook.get('sections', []):
                if section_id and section['id'] != section_id:
                    continue

                for page in section.get('pages', []):
                    page['section_id'] = section['id']
                    page['section_name'] = section['name']
                    page['notebook_id'] = notebook['id']
                    page['notebook_name'] = notebook['name']
                    pages.append(page)

        return pages

    def find_notebook(self, name: str) -> Optional[Dict[str, str]]:
        """
        Find notebook by name.

        Args:
            name: Notebook name

        Returns:
            dict or None: Notebook info if found
        """
        notebooks = self.get_notebooks()
        for notebook in notebooks:
            if notebook['name'].lower() == name.lower():
                return notebook
        return None

    def find_section(self, notebook_name: str, section_name: str) -> Optional[Dict[str, str]]:
        """
        Find section by notebook and section name.

        Args:
            notebook_name: Notebook name
            section_name: Section name

        Returns:
            dict or None: Section info if found
        """
        notebook = self.find_notebook(notebook_name)
        if not notebook:
            return None

        sections = self.get_sections(notebook['id'])
        for section in sections:
            if section['name'].lower() == section_name.lower():
                return section
        return None

    def find_page(self, notebook_name: str, section_name: str, page_title: str) -> Optional[Dict[str, str]]:
        """
        Find page by notebook, section, and title.

        Args:
            notebook_name: Notebook name
            section_name: Section name
            page_title: Page title

        Returns:
            dict or None: Page info if found
        """
        section = self.find_section(notebook_name, section_name)
        if not section:
            return None

        pages = self.get_pages(section['id'])
        for page in pages:
            if page['title'].lower() == page_title.lower():
                return page
        return None

    # Page Content Methods

    def get_page_content(self, page_id: str, format: str = "simple") -> str:
        """
        Get page content.

        Args:
            page_id: Page ID
            format: Content format ('xml', 'simple', 'text')

        Returns:
            str: Page content
        """
        self._ensure_connected()

        try:
            # Get page content XML
            xml_content = self.onenote.GetPageContent(page_id)

            if format == "xml":
                return xml_content
            elif format == "simple":
                return self._parse_page_content_simple(xml_content)
            elif format == "text":
                return self._extract_text_from_page(xml_content)
            else:
                return xml_content

        except Exception as e:
            self.logger.error(f"Failed to get page content: {e}")
            raise

    def create_page(
        self,
        section_id: str,
        title: str,
        content: str = "",
        datetime_str: str = None
    ) -> str:
        """
        Create a new page.

        Args:
            section_id: Section ID to create page in
            title: Page title
            content: Page content (HTML or plain text)
            datetime_str: Optional datetime for page

        Returns:
            str: New page ID
        """
        self._ensure_connected()

        try:
            # Generate page XML
            page_xml = self._build_page_xml(title, content, datetime_str)

            # Create page
            page_id = ""
            self.onenote.CreateNewPage(section_id, page_id, 0)  # 0 = default position

            # Get the newly created page ID
            pages = self.get_pages(section_id)
            if pages:
                page_id = pages[0]['id']  # Most recent page

            # Update content if provided
            if content:
                self.update_page_content(page_id, title, content)

            self.logger.info(f"Created page: {title}")
            return page_id

        except Exception as e:
            self.logger.error(f"Failed to create page: {e}")
            raise

    def update_page_content(
        self,
        page_id: str,
        title: str = None,
        content: str = ""
    ) -> bool:
        """
        Update page content.

        Args:
            page_id: Page ID
            title: Optional new title
            content: New content (HTML or plain text)

        Returns:
            bool: True if successful
        """
        self._ensure_connected()

        try:
            # Get current page content
            current_xml = self.onenote.GetPageContent(page_id)
            root = ET.fromstring(current_xml)

            # Update title if provided
            if title:
                ns = {'one': self.NS}
                title_elem = root.find('.//one:Title', ns)
                if title_elem is not None:
                    oe_elem = title_elem.find('one:OE', ns)
                    if oe_elem is not None:
                        t_elem = oe_elem.find('one:T', ns)
                        if t_elem is not None:
                            t_elem.text = title

            # Build content XML and append
            if content:
                content_xml = self._build_content_xml(content)
                # Append to outline
                outline = root.find('.//{%s}Outline' % self.NS)
                if outline is not None:
                    content_elem = ET.fromstring(content_xml)
                    outline.append(content_elem)

            # Update page
            updated_xml = ET.tostring(root, encoding='unicode')
            self.onenote.UpdatePageContent(updated_xml)

            self.logger.info(f"Updated page: {page_id}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to update page: {e}")
            raise

    def delete_page(self, page_id: str) -> bool:
        """
        Delete a page.

        Args:
            page_id: Page ID

        Returns:
            bool: True if successful
        """
        self._ensure_connected()

        try:
            self.onenote.DeletePageContent(page_id)
            self.logger.info(f"Deleted page: {page_id}")
            return True
        except Exception as e:
            self.logger.error(f"Failed to delete page: {e}")
            raise

    # Search Methods

    def search_pages(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        """
        Search for pages containing text.

        Args:
            query: Search query
            max_results: Maximum results to return

        Returns:
            list: List of matching pages with snippets
        """
        self._ensure_connected()

        try:
            # Get all pages
            all_pages = self.get_pages()

            # Search through pages
            results = []
            for page in all_pages[:max_results]:
                try:
                    content = self.get_page_content(page['id'], format='text')
                    if query.lower() in content.lower():
                        # Find snippet around match
                        idx = content.lower().index(query.lower())
                        start = max(0, idx - 50)
                        end = min(len(content), idx + len(query) + 50)
                        snippet = content[start:end].strip()

                        results.append({
                            **page,
                            'snippet': snippet,
                            'match_position': idx
                        })
                except Exception as e:
                    self.logger.warning(f"Error searching page {page['id']}: {e}")
                    continue

            self.logger.info(f"Search found {len(results)} results for '{query}'")
            return results

        except Exception as e:
            self.logger.error(f"Search failed: {e}")
            raise

    # Helper Methods

    def _parse_hierarchy_xml(self, xml_string: str) -> Dict[str, Any]:
        """Parse OneNote hierarchy XML."""
        try:
            root = ET.fromstring(xml_string)
            ns = {'one': self.NS}

            notebooks = []
            for nb in root.findall('one:Notebook', ns):
                notebook = {
                    'id': nb.get('ID'),
                    'name': nb.get('name'),
                    'path': nb.get('path', ''),
                    'sections': []
                }

                for sec in nb.findall('.//one:Section', ns):
                    section = {
                        'id': sec.get('ID'),
                        'name': sec.get('name'),
                        'path': sec.get('path', ''),
                        'pages': []
                    }

                    for pg in sec.findall('one:Page', ns):
                        page = {
                            'id': pg.get('ID'),
                            'title': pg.get('name', 'Untitled'),
                            'datetime': pg.get('dateTime', ''),
                            'last_modified': pg.get('lastModifiedTime', '')
                        }
                        section['pages'].append(page)

                    notebook['sections'].append(section)

                notebooks.append(notebook)

            return {'notebooks': notebooks}

        except Exception as e:
            self.logger.error(f"Failed to parse hierarchy XML: {e}")
            raise

    def _parse_page_content_simple(self, xml_string: str) -> str:
        """Parse page content XML to simple format."""
        try:
            root = ET.fromstring(xml_string)
            ns = {'one': self.NS}

            content_parts = []

            # Get title
            title_elem = root.find('.//one:Title', ns)
            if title_elem is not None:
                title_text = self._get_text_from_element(title_elem, ns)
                if title_text:
                    content_parts.append(f"# {title_text}\n")

            # Get content
            for outline in root.findall('.//one:Outline', ns):
                text = self._get_text_from_element(outline, ns)
                if text:
                    content_parts.append(text)

            return "\n\n".join(content_parts)

        except Exception as e:
            self.logger.error(f"Failed to parse page content: {e}")
            return ""

    def _extract_text_from_page(self, xml_string: str) -> str:
        """Extract plain text from page XML."""
        try:
            root = ET.fromstring(xml_string)
            ns = {'one': self.NS}

            text_parts = []
            for t_elem in root.findall('.//one:T', ns):
                if t_elem.text:
                    text_parts.append(t_elem.text.strip())

            return " ".join(text_parts)

        except Exception as e:
            self.logger.error(f"Failed to extract text: {e}")
            return ""

    def _get_text_from_element(self, element: ET.Element, ns: dict) -> str:
        """Recursively get text from XML element."""
        text_parts = []
        for t_elem in element.findall('.//one:T', ns):
            if t_elem.text:
                text_parts.append(t_elem.text.strip())
        return " ".join(text_parts)

    def _build_page_xml(self, title: str, content: str, datetime_str: str = None) -> str:
        """Build XML for new page."""
        if datetime_str is None:
            datetime_str = datetime.now().strftime("%Y-%m-%dT%H:%M:%S.000Z")

        xml = f'''<?xml version="1.0"?>
<one:Page xmlns:one="{self.NS}" ID="{{00000000-0000-0000-0000-000000000000}}" dateTime="{datetime_str}">
    <one:Title>
        <one:OE>
            <one:T><![CDATA[{title}]]></one:T>
        </one:OE>
    </one:Title>
    <one:Outline>
        {self._build_content_xml(content)}
    </one:Outline>
</one:Page>'''
        return xml

    def _build_content_xml(self, content: str) -> str:
        """Build XML for page content."""
        # Simple text content for now
        lines = content.split('\n')
        xml_parts = []

        for line in lines:
            if line.strip():
                xml_parts.append(f'<one:OE><one:T><![CDATA[{line}]]></one:T></one:OE>')

        return '\n'.join(xml_parts)

    def __repr__(self) -> str:
        status = "connected" if self._connected else "disconnected"
        return f"OneNoteCOMClient({status})"
