"""
OneNote Helpers

Helper class for OneNote knowledge management tasks.
"""

from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional

from core.logging_config import get_logger

logger = get_logger(__name__)


class OneNoteHelper:
    """Helper for OneNote knowledge management tasks."""

    def __init__(
        self,
        default_notebook: Optional[str] = None,
        default_section: Optional[str] = None
    ) -> None:
        """
        Initialize OneNote helper.

        Args:
            default_notebook: Default notebook name.
            default_section: Default section name.
        """
        from modules.onenote import OneNoteManager, OneNoteContentBuilder, TemplateBuilder

        self.manager = OneNoteManager()
        self.manager.configure(
            default_notebook=default_notebook,
            default_section=default_section
        )
        self.manager.validate()

        self.content_builder = OneNoteContentBuilder
        self.templates = TemplateBuilder

    def create_page(
        self,
        notebook: str,
        section: str,
        title: str,
        content: str = ""
    ) -> str:
        """
        Create a simple OneNote page.

        Args:
            notebook: Notebook name.
            section: Section name.
            title: Page title.
            content: Page content (simple text).

        Returns:
            New page ID.

        Example:
            onenote = OneNoteHelper()
            page_id = onenote.create_page(
                notebook='Project Knowledge',
                section='Requirements',
                title='Q4 Requirements',
                content='Key requirements for Q4...'
            )
        """
        logger.info(f"Creating OneNote page: {title}")

        page_id = self.manager.create_page(
            notebook=notebook,
            section=section,
            title=title,
            content=content
        )

        logger.info(f"OneNote page created: {page_id}")
        return page_id

    def get_page_content(
        self,
        notebook: str,
        section: str,
        title: str,
        format: str = "text"
    ) -> Optional[str]:
        """
        Get content from a OneNote page.

        Args:
            notebook: Notebook name.
            section: Section name.
            title: Page title.
            format: Content format ('text', 'simple', 'xml').

        Returns:
            Page content or None if not found.

        Example:
            onenote = OneNoteHelper()
            content = onenote.get_page_content(
                notebook='Project Knowledge',
                section='Requirements',
                title='Q4 Requirements'
            )
            print(content)
        """
        logger.info(f"Getting OneNote page content: {title}")

        content = self.manager.get_page(
            notebook=notebook,
            section=section,
            title=title,
            format=format
        )

        if content:
            logger.info(f"Retrieved {len(content)} characters from page")
        else:
            logger.warning(f"Page '{title}' not found")

        return content

    def search_for_facts(
        self,
        query: str,
        notebook: Optional[str] = None,
        max_results: int = 10
    ) -> List[Dict[str, Any]]:
        """
        Search OneNote for project facts/information.

        Args:
            query: Search query.
            notebook: Optional notebook to search in.
            max_results: Maximum results to return.

        Returns:
            Matching pages with snippets.

        Example:
            onenote = OneNoteHelper()
            results = onenote.search_for_facts(
                query='approved requirements',
                notebook='Project Knowledge'
            )

            for result in results:
                print(f"{result['title']}: {result['snippet']}")
        """
        logger.info(f"Searching OneNote for: '{query}'")

        results = self.manager.search(
            query=query,
            notebook=notebook,
            max_results=max_results
        )

        logger.info(f"Found {len(results)} results")
        return results

    def extract_facts_for_report(
        self,
        notebook: str,
        section: str,
        page_title: str
    ) -> Dict[str, Any]:
        """
        Extract structured facts from OneNote page for use in reports.

        Args:
            notebook: Notebook name.
            section: Section name.
            page_title: Page title.

        Returns:
            Extracted facts and tables.

        Example:
            onenote = OneNoteHelper()
            facts = onenote.extract_facts_for_report(
                notebook='Project Knowledge',
                section='Status',
                page_title='Q4 Metrics'
            )

            # Use in Excel report
            report = ReportBuilder()
            report.add_data_sheet('Project Facts', facts['tables'][0])
        """
        logger.info(f"Extracting facts from OneNote page: {page_title}")

        # Get page content
        content = self.manager.get_page(
            notebook=notebook,
            section=section,
            title=page_title,
            format='text'
        )

        # Extract tables
        tables = self.manager.extract_tables_from_page(
            notebook=notebook,
            section=section,
            title=page_title
        )

        facts = {
            'content': content,
            'tables': tables,
            'source': f"{notebook}/{section}/{page_title}",
            'extracted_at': datetime.now().isoformat()
        }

        logger.info(f"Extracted {len(tables)} tables from page")
        return facts

    def create_meeting_minutes(
        self,
        notebook: str,
        section: str,
        meeting_title: str,
        date: str,
        attendees: List[str],
        agenda: List[str],
        notes: str = "",
        action_items: Optional[List[Dict[str, str]]] = None
    ) -> str:
        """
        Create meeting minutes page from template.

        Args:
            notebook: Notebook name.
            section: Section name.
            meeting_title: Meeting title.
            date: Meeting date.
            attendees: List of attendees.
            agenda: Agenda items.
            notes: Meeting notes.
            action_items: Action items with 'task', 'owner', 'due' keys.

        Returns:
            New page ID.

        Example:
            onenote = OneNoteHelper()
            page_id = onenote.create_meeting_minutes(
                notebook='Meetings',
                section='Sprint Planning',
                meeting_title='Sprint 42 Planning',
                date='2024-12-09',
                attendees=['Alice', 'Bob', 'Carol'],
                agenda=['Review backlog', 'Estimate tasks', 'Plan sprint'],
                action_items=[
                    {'task': 'Update design docs', 'owner': 'Alice', 'due': '2024-12-12'}
                ]
            )
        """
        logger.info(f"Creating meeting minutes: {meeting_title}")

        page_id = self.manager.create_meeting_minutes(
            notebook=notebook,
            section=section,
            meeting_title=meeting_title,
            date=date,
            attendees=attendees,
            agenda=agenda,
            notes=notes,
            action_items=action_items
        )

        logger.info(f"Meeting minutes created: {page_id}")
        return page_id

    def create_meeting_minutes_from_outlook(
        self,
        notebook: str,
        section: str,
        meeting_subject: str
    ) -> str:
        """
        Create meeting minutes from Outlook meeting details.

        Args:
            notebook: Notebook name.
            section: Section name.
            meeting_subject: Outlook meeting subject to find.

        Returns:
            New page ID.

        Example:
            onenote = OneNoteHelper()
            page_id = onenote.create_meeting_minutes_from_outlook(
                notebook='Meetings',
                section='Sprint Planning',
                meeting_subject='Sprint 42 Planning Meeting'
            )
        """
        from modules.outlook_automation import MeetingManager

        logger.info(f"Creating meeting minutes from Outlook: {meeting_subject}")

        # Get meeting details from Outlook
        meeting_mgr = MeetingManager()
        meeting_mgr.connect()
        meeting = meeting_mgr.get_meeting_by_subject(meeting_subject)

        if not meeting:
            raise ValueError(f"Meeting '{meeting_subject}' not found in Outlook")

        # Extract details
        attendees = [a.get('name', '') for a in meeting.get('attendees', [])]
        date = meeting.get('start_time', datetime.now().strftime('%Y-%m-%d'))
        body = meeting.get('body', '')

        # Parse agenda from body (simple approach)
        agenda: List[str] = []
        for line in body.split('\n'):
            line = line.strip()
            if line and (line.startswith('-') or line.startswith('•')):
                agenda.append(line.lstrip('-•').strip())

        if not agenda:
            agenda = ['Review agenda items']

        # Create page
        page_id = self.create_meeting_minutes(
            notebook=notebook,
            section=section,
            meeting_title=meeting.get('subject', meeting_subject),
            date=date,
            attendees=attendees,
            agenda=agenda,
            notes=""
        )

        logger.info("Meeting minutes created from Outlook meeting")
        return page_id

    def create_project_status_page(
        self,
        notebook: str,
        section: str,
        project_name: str,
        status: str,
        progress: int,
        milestones: List[Dict[str, str]],
        issues: Optional[List[str]] = None,
        next_steps: Optional[List[str]] = None
    ) -> str:
        """
        Create project status page from template.

        Args:
            notebook: Notebook name.
            section: Section name.
            project_name: Project name.
            status: Status (On Track, At Risk, Delayed).
            progress: Progress percentage (0-100).
            milestones: List with 'name', 'due', 'status' keys.
            issues: List of issues.
            next_steps: List of next steps.

        Returns:
            New page ID.

        Example:
            onenote = OneNoteHelper()
            page_id = onenote.create_project_status_page(
                notebook='Projects',
                section='Q4 2024',
                project_name='Homepage Redesign',
                status='On Track',
                progress=75,
                milestones=[
                    {'name': 'Design Complete', 'due': '2024-12-01', 'status': 'Done'},
                    {'name': 'Development', 'due': '2024-12-15', 'status': 'In Progress'}
                ],
                next_steps=['Complete testing', 'Deploy to production']
            )
        """
        logger.info(f"Creating project status page: {project_name}")

        page_id = self.manager.create_project_status(
            notebook=notebook,
            section=section,
            project_name=project_name,
            status=status,
            progress=progress,
            milestones=milestones,
            issues=issues,
            next_steps=next_steps
        )

        logger.info(f"Project status page created: {page_id}")
        return page_id

    def create_research_notes(
        self,
        notebook: str,
        section: str,
        topic: str,
        source: str,
        key_findings: List[str],
        details: str = "",
        tags: Optional[List[str]] = None
    ) -> str:
        """
        Create research notes page.

        Args:
            notebook: Notebook name.
            section: Section name.
            topic: Research topic.
            source: Information source.
            key_findings: List of key findings.
            details: Additional details.
            tags: Tags for categorization.

        Returns:
            New page ID.

        Example:
            onenote = OneNoteHelper()
            page_id = onenote.create_research_notes(
                notebook='Research',
                section='Competitors',
                topic='Competitor Homepage Analysis',
                source='Web research - 2024-12-09',
                key_findings=[
                    'Average load time: 2.3 seconds',
                    'Mobile-first design approach',
                    'Focus on video content'
                ],
                tags=['homepage', 'ux', 'performance']
            )
        """
        logger.info(f"Creating research notes: {topic}")

        page_id = self.manager.create_research_notes(
            notebook=notebook,
            section=section,
            topic=topic,
            source=source,
            key_findings=key_findings,
            details=details,
            tags=tags
        )

        logger.info(f"Research notes created: {page_id}")
        return page_id

    def aggregate_research_to_central_notebook(
        self,
        source_files: List[str],
        central_notebook: str,
        section: str,
        topic: str
    ) -> str:
        """
        Aggregate research from various files into central OneNote repository.

        Args:
            source_files: List of file paths to aggregate (txt, pdf, md, etc.).
            central_notebook: Central notebook name.
            section: Section name.
            topic: Research topic.

        Returns:
            New page ID.

        Example:
            onenote = OneNoteHelper()
            page_id = onenote.aggregate_research_to_central_notebook(
                source_files=[
                    'C:/Research/analysis1.txt',
                    'C:/Research/analysis2.txt'
                ],
                central_notebook='Project Knowledge',
                section='Research',
                topic='Market Analysis'
            )
        """
        logger.info(f"Aggregating research from {len(source_files)} files")

        # Read all source files
        all_content: List[str] = []
        key_findings: List[str] = []

        for file_path in source_files:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    all_content.append(f"## {Path(file_path).name}\n\n{content}\n")

                    # Extract first few lines as findings (simple approach)
                    lines = [line.strip() for line in content.split('\n') if line.strip()]
                    key_findings.extend(lines[:3])

                logger.info(f"Read: {Path(file_path).name}")
            except Exception as e:
                logger.warning(f"Failed to read {file_path}: {e}")

        # Combine content
        combined_content = "\n\n".join(all_content)

        # Create research notes page
        page_id = self.create_research_notes(
            notebook=central_notebook,
            section=section,
            topic=topic,
            source=f"Aggregated from {len(source_files)} files",
            key_findings=key_findings[:10],  # Top 10 findings
            details=combined_content,
            tags=['aggregated', 'research']
        )

        logger.info(f"Aggregated research into OneNote page: {page_id}")
        return page_id

    def list_notebooks(self) -> List[Dict[str, str]]:
        """
        List all available OneNote notebooks.

        Returns:
            Notebooks with id, name, path.

        Example:
            onenote = OneNoteHelper()
            notebooks = onenote.list_notebooks()
            for nb in notebooks:
                print(f"Notebook: {nb['name']}")
        """
        return self.manager.list_notebooks()

    def list_sections(self, notebook: str) -> List[Dict[str, str]]:
        """
        List sections in a notebook.

        Args:
            notebook: Notebook name.

        Returns:
            Sections with id, name, path.

        Example:
            onenote = OneNoteHelper()
            sections = onenote.list_sections('Project Knowledge')
            for sec in sections:
                print(f"Section: {sec['name']}")
        """
        return self.manager.list_sections(notebook)

    def list_pages(self, notebook: str, section: str) -> List[Dict[str, str]]:
        """
        List pages in a section.

        Args:
            notebook: Notebook name.
            section: Section name.

        Returns:
            Pages with id, title, datetime.

        Example:
            onenote = OneNoteHelper()
            pages = onenote.list_pages('Project Knowledge', 'Requirements')
            for page in pages:
                print(f"Page: {page['title']}")
        """
        return self.manager.list_pages(notebook, section)
