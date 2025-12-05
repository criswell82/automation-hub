"""
OneNote Manager Module.
High-level API for OneNote automation using COM interface.
"""

import sys
from pathlib import Path
from typing import Dict, List, Optional, Any

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from base_module import BaseModule
from onenote.com_client import OneNoteCOMClient
from onenote.content_formatter import OneNoteContentBuilder, TemplateBuilder


class OneNoteManager(BaseModule):
    """
    High-level OneNote automation manager.
    Provides user-friendly API for common OneNote operations.
    """

    def __init__(self) -> None:
        """Initialize OneNote Manager."""
        super().__init__(
            name="OneNoteManager",
            description="OneNote automation using COM interface",
            version="1.0.0",
            category="microsoft_office"
        )

        self.client: Optional[OneNoteCOMClient] = None
        self._default_notebook = None
        self._default_section = None

    def configure(
        self,
        default_notebook: Optional[str] = None,
        default_section: Optional[str] = None,
        **kwargs: Any
    ) -> bool:
        """
        Configure OneNote manager.

        Args:
            default_notebook: Default notebook name
            default_section: Default section name
            **kwargs: Additional configuration

        Returns:
            bool: True if successful
        """
        try:
            self.log_info("Configuring OneNote manager...")

            # Store defaults
            self._default_notebook = default_notebook
            self._default_section = default_section

            # Store in module config
            self.set_config('default_notebook', default_notebook)
            self.set_config('default_section', default_section)

            self.log_info("Configuration complete")
            return True

        except Exception as e:
            self.log_error(f"Configuration failed: {e}")
            return False

    def validate(self) -> bool:
        """
        Validate configuration and OneNote connection.

        Returns:
            bool: True if valid
        """
        try:
            self.log_info("Validating OneNote connection...")

            # Create and connect client
            self.client = OneNoteCOMClient()
            self.client.connect()

            # Verify we can access notebooks
            notebooks = self.client.get_notebooks()
            self.log_info(f"Found {len(notebooks)} notebooks")

            # Validate default notebook exists if specified
            if self._default_notebook:
                nb = self.client.find_notebook(self._default_notebook)
                if not nb:
                    self.log_warning(f"Default notebook '{self._default_notebook}' not found")
                    return False

            self.log_info("Validation successful")
            return True

        except Exception as e:
            self.log_error(f"Validation failed: {e}")
            return False

    def execute(self) -> Dict[str, Any]:
        """
        Execute OneNote operation.
        This is a generic execution - specific operations should use dedicated methods.

        Returns:
            dict: Execution result
        """
        try:
            # Get operation from config
            operation = self.get_config('operation')

            if operation == 'list_notebooks':
                notebooks = self.list_notebooks()
                return self.success_result(notebooks, "Listed notebooks")

            elif operation == 'list_sections':
                sections = self.list_sections()
                return self.success_result(sections, "Listed sections")

            elif operation == 'list_pages':
                pages = self.list_pages()
                return self.success_result(pages, "Listed pages")

            elif operation == 'create_page':
                page_id = self.create_page(
                    notebook=self.get_config('notebook'),
                    section=self.get_config('section'),
                    title=self.get_config('title'),
                    content=self.get_config('content', '')
                )
                return self.success_result({'page_id': page_id}, "Page created")

            elif operation == 'search':
                results = self.search(self.get_config('query'))
                return self.success_result(results, f"Found {len(results)} results")

            else:
                return self.error_result(f"Unknown operation: {operation}")

        except Exception as e:
            self.log_error(f"Execution failed: {e}")
            raise

    # High-level API Methods

    def list_notebooks(self) -> List[Dict[str, str]]:
        """
        List all notebooks.

        Returns:
            list: Notebooks with id, name, path
        """
        self.log_info("Listing notebooks...")
        notebooks = self.client.get_notebooks()
        self.log_info(f"Found {len(notebooks)} notebooks")
        return notebooks

    def list_sections(self, notebook: str = None) -> List[Dict[str, str]]:
        """
        List sections in a notebook.

        Args:
            notebook: Notebook name (defaults to configured default)

        Returns:
            list: Sections with id, name, path
        """
        notebook = notebook or self._default_notebook

        if notebook:
            nb = self.client.find_notebook(notebook)
            if nb:
                sections = self.client.get_sections(nb['id'])
            else:
                self.log_warning(f"Notebook '{notebook}' not found")
                sections = []
        else:
            sections = self.client.get_sections()

        self.log_info(f"Found {len(sections)} sections")
        return sections

    def list_pages(
        self,
        notebook: str = None,
        section: str = None
    ) -> List[Dict[str, str]]:
        """
        List pages in a section.

        Args:
            notebook: Notebook name
            section: Section name

        Returns:
            list: Pages with id, title, datetime
        """
        notebook = notebook or self._default_notebook
        section = section or self._default_section

        if notebook and section:
            sec = self.client.find_section(notebook, section)
            if sec:
                pages = self.client.get_pages(sec['id'])
            else:
                self.log_warning(f"Section '{section}' not found in '{notebook}'")
                pages = []
        else:
            pages = self.client.get_pages()

        self.log_info(f"Found {len(pages)} pages")
        return pages

    def create_page(
        self,
        notebook: str,
        section: str,
        title: str,
        content: Any = None,
        template: str = None
    ) -> str:
        """
        Create a new OneNote page.

        Args:
            notebook: Notebook name
            section: Section name
            title: Page title
            content: Page content (string, OneNoteContentBuilder, or dict for template)
            template: Template name (meeting_minutes, project_status, research_notes, decision_log)

        Returns:
            str: New page ID
        """
        self.log_info(f"Creating page '{title}' in {notebook}/{section}...")

        # Find section
        sec = self.client.find_section(notebook, section)
        if not sec:
            # Try to create section if it doesn't exist
            self.log_warning(f"Section '{section}' not found, attempting to create...")
            # For now, raise error - section creation requires additional logic
            raise ValueError(f"Section '{section}' not found in notebook '{notebook}'")

        # Build content
        if isinstance(content, OneNoteContentBuilder):
            content_text = content.build_simple()
        elif isinstance(content, dict) and template:
            content_text = self._build_from_template(template, **content).build_simple()
        elif content is None:
            content_text = ""
        else:
            content_text = str(content)

        # Create page
        page_id = self.client.create_page(
            section_id=sec['id'],
            title=title,
            content=content_text
        )

        self.log_info(f"Page created with ID: {page_id}")
        return page_id

    def get_page(
        self,
        notebook: str,
        section: str,
        title: str,
        format: str = "simple"
    ) -> Optional[str]:
        """
        Get page content.

        Args:
            notebook: Notebook name
            section: Section name
            title: Page title
            format: Content format ('simple', 'text', 'xml')

        Returns:
            str or None: Page content
        """
        self.log_info(f"Getting page '{title}' from {notebook}/{section}...")

        # Find page
        page = self.client.find_page(notebook, section, title)
        if not page:
            self.log_warning(f"Page '{title}' not found")
            return None

        # Get content
        content = self.client.get_page_content(page['id'], format=format)
        self.log_info(f"Retrieved page content ({len(content)} chars)")
        return content

    def update_page(
        self,
        notebook: str,
        section: str,
        title: str,
        content: Any,
        new_title: str = None
    ) -> bool:
        """
        Update page content.

        Args:
            notebook: Notebook name
            section: Section name
            title: Current page title
            content: New content
            new_title: New title (optional)

        Returns:
            bool: True if successful
        """
        self.log_info(f"Updating page '{title}' in {notebook}/{section}...")

        # Find page
        page = self.client.find_page(notebook, section, title)
        if not page:
            self.log_warning(f"Page '{title}' not found")
            return False

        # Build content
        if isinstance(content, OneNoteContentBuilder):
            content_text = content.build_simple()
        else:
            content_text = str(content)

        # Update
        success = self.client.update_page_content(
            page_id=page['id'],
            title=new_title,
            content=content_text
        )

        if success:
            self.log_info("Page updated successfully")
        return success

    def delete_page(
        self,
        notebook: str,
        section: str,
        title: str
    ) -> bool:
        """
        Delete a page.

        Args:
            notebook: Notebook name
            section: Section name
            title: Page title

        Returns:
            bool: True if successful
        """
        self.log_info(f"Deleting page '{title}' from {notebook}/{section}...")

        # Find page
        page = self.client.find_page(notebook, section, title)
        if not page:
            self.log_warning(f"Page '{title}' not found")
            return False

        # Delete
        success = self.client.delete_page(page['id'])

        if success:
            self.log_info("Page deleted successfully")
        return success

    def search(
        self,
        query: str,
        notebook: str = None,
        max_results: int = 50
    ) -> List[Dict[str, Any]]:
        """
        Search for pages containing text.

        Args:
            query: Search query
            notebook: Optional notebook to search in
            max_results: Maximum results

        Returns:
            list: Matching pages with snippets
        """
        self.log_info(f"Searching for '{query}'...")

        results = self.client.search_pages(query, max_results=max_results)

        # Filter by notebook if specified
        if notebook:
            nb = self.client.find_notebook(notebook)
            if nb:
                results = [r for r in results if r.get('notebook_id') == nb['id']]

        self.log_info(f"Found {len(results)} results")
        return results

    def extract_tables_from_page(
        self,
        notebook: str,
        section: str,
        title: str
    ) -> List[List[List[str]]]:
        """
        Extract tables from a page.

        Args:
            notebook: Notebook name
            section: Section name
            title: Page title

        Returns:
            list: List of tables (each table is a list of rows)
        """
        self.log_info(f"Extracting tables from '{title}'...")

        content = self.get_page(notebook, section, title, format='text')
        if not content:
            return []

        # Simple table extraction (looking for | delimited content)
        tables = []
        current_table = []
        in_table = False

        for line in content.split('\n'):
            if '|' in line and len(line.split('|')) > 2:
                # Likely a table row
                row = [cell.strip() for cell in line.split('|') if cell.strip()]
                current_table.append(row)
                in_table = True
            elif in_table and current_table:
                # End of table
                tables.append(current_table)
                current_table = []
                in_table = False

        # Add last table if exists
        if current_table:
            tables.append(current_table)

        self.log_info(f"Extracted {len(tables)} tables")
        return tables

    # Template Helpers

    def create_meeting_minutes(
        self,
        notebook: str,
        section: str,
        meeting_title: str,
        date: str,
        attendees: List[str],
        agenda: List[str],
        notes: str = "",
        action_items: List[Dict[str, str]] = None
    ) -> str:
        """
        Create meeting minutes page.

        Args:
            notebook: Notebook name
            section: Section name
            meeting_title: Meeting title
            date: Meeting date
            attendees: Attendee list
            agenda: Agenda items
            notes: Meeting notes
            action_items: Action items

        Returns:
            str: New page ID
        """
        builder = TemplateBuilder.meeting_minutes(
            meeting_title=meeting_title,
            date=date,
            attendees=attendees,
            agenda_items=agenda,
            notes=notes,
            action_items=action_items
        )

        return self.create_page(
            notebook=notebook,
            section=section,
            title=meeting_title,
            content=builder
        )

    def create_project_status(
        self,
        notebook: str,
        section: str,
        project_name: str,
        status: str,
        progress: int,
        milestones: List[Dict[str, str]],
        issues: List[str] = None,
        next_steps: List[str] = None
    ) -> str:
        """
        Create project status page.

        Args:
            notebook: Notebook name
            section: Section name
            project_name: Project name
            status: Overall status
            progress: Progress percentage
            milestones: Milestone list
            issues: Issue list
            next_steps: Next steps

        Returns:
            str: New page ID
        """
        builder = TemplateBuilder.project_status(
            project_name=project_name,
            status=status,
            progress=progress,
            milestones=milestones,
            issues=issues,
            next_steps=next_steps
        )

        return self.create_page(
            notebook=notebook,
            section=section,
            title=f"{project_name} - Status Report",
            content=builder
        )

    def create_research_notes(
        self,
        notebook: str,
        section: str,
        topic: str,
        source: str,
        key_findings: List[str],
        details: str = "",
        tags: List[str] = None
    ) -> str:
        """
        Create research notes page.

        Args:
            notebook: Notebook name
            section: Section name
            topic: Research topic
            source: Information source
            key_findings: Key findings
            details: Additional details
            tags: Tags

        Returns:
            str: New page ID
        """
        builder = TemplateBuilder.research_notes(
            topic=topic,
            source=source,
            key_findings=key_findings,
            details=details,
            tags=tags
        )

        return self.create_page(
            notebook=notebook,
            section=section,
            title=f"Research: {topic}",
            content=builder
        )

    def _build_from_template(self, template: str, **kwargs) -> OneNoteContentBuilder:
        """Build content from template."""
        if template == 'meeting_minutes':
            return TemplateBuilder.meeting_minutes(**kwargs)
        elif template == 'project_status':
            return TemplateBuilder.project_status(**kwargs)
        elif template == 'research_notes':
            return TemplateBuilder.research_notes(**kwargs)
        elif template == 'decision_log':
            return TemplateBuilder.decision_log(**kwargs)
        else:
            raise ValueError(f"Unknown template: {template}")

    def __repr__(self) -> str:
        return f"OneNoteManager(status='{self.status}', notebooks={self._default_notebook})"
