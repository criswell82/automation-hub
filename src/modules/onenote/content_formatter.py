"""
OneNote Content Formatter.
Provides helpers for building rich OneNote content (text, tables, lists, images).
"""

from typing import List, Dict, Any
from datetime import datetime


class OneNoteContentBuilder:
    """Builder for creating OneNote page content."""

    NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

    def __init__(self, title: str = ""):
        """
        Initialize content builder.

        Args:
            title: Page title
        """
        self.title = title
        self.content_parts = []

    def add_heading(self, text: str, level: int = 1) -> 'OneNoteContentBuilder':
        """
        Add heading.

        Args:
            text: Heading text
            level: Heading level (1-6)

        Returns:
            self: For method chaining
        """
        # OneNote headings are represented with style attributes
        size = max(12, 24 - (level * 2))  # Decrease size based on level
        self.content_parts.append({
            'type': 'heading',
            'text': text,
            'level': level,
            'size': size
        })
        return self

    def add_text(self, text: str, bold: bool = False, italic: bool = False) -> 'OneNoteContentBuilder':
        """
        Add text paragraph.

        Args:
            text: Text content
            bold: Make text bold
            italic: Make text italic

        Returns:
            self: For method chaining
        """
        self.content_parts.append({
            'type': 'text',
            'text': text,
            'bold': bold,
            'italic': italic
        })
        return self

    def add_bullet_list(self, items: List[str]) -> 'OneNoteContentBuilder':
        """
        Add bullet list.

        Args:
            items: List items

        Returns:
            self: For method chaining
        """
        self.content_parts.append({
            'type': 'bullet_list',
            'items': items
        })
        return self

    def add_numbered_list(self, items: List[str]) -> 'OneNoteContentBuilder':
        """
        Add numbered list.

        Args:
            items: List items

        Returns:
            self: For method chaining
        """
        self.content_parts.append({
            'type': 'numbered_list',
            'items': items
        })
        return self

    def add_table(self, headers: List[str], rows: List[List[str]]) -> 'OneNoteContentBuilder':
        """
        Add table.

        Args:
            headers: Table headers
            rows: Table rows

        Returns:
            self: For method chaining
        """
        self.content_parts.append({
            'type': 'table',
            'headers': headers,
            'rows': rows
        })
        return self

    def add_code_block(self, code: str, language: str = "") -> 'OneNoteContentBuilder':
        """
        Add code block.

        Args:
            code: Code content
            language: Programming language (optional)

        Returns:
            self: For method chaining
        """
        self.content_parts.append({
            'type': 'code',
            'text': code,
            'language': language
        })
        return self

    def add_divider(self) -> 'OneNoteContentBuilder':
        """
        Add horizontal divider.

        Returns:
            self: For method chaining
        """
        self.content_parts.append({'type': 'divider'})
        return self

    def add_blank_line(self) -> 'OneNoteContentBuilder':
        """
        Add blank line.

        Returns:
            self: For method chaining
        """
        self.content_parts.append({'type': 'blank'})
        return self

    def build_simple(self) -> str:
        """
        Build simple text content.

        Returns:
            str: Simple text content
        """
        lines = []

        if self.title:
            lines.append(f"# {self.title}")
            lines.append("")

        for part in self.content_parts:
            part_type = part['type']

            if part_type == 'heading':
                prefix = '#' * part['level']
                lines.append(f"{prefix} {part['text']}")
                lines.append("")

            elif part_type == 'text':
                text = part['text']
                if part.get('bold'):
                    text = f"**{text}**"
                if part.get('italic'):
                    text = f"*{text}*"
                lines.append(text)
                lines.append("")

            elif part_type == 'bullet_list':
                for item in part['items']:
                    lines.append(f"• {item}")
                lines.append("")

            elif part_type == 'numbered_list':
                for i, item in enumerate(part['items'], 1):
                    lines.append(f"{i}. {item}")
                lines.append("")

            elif part_type == 'table':
                # Simple ASCII table
                headers = part['headers']
                rows = part['rows']

                # Calculate column widths
                col_widths = [len(h) for h in headers]
                for row in rows:
                    for i, cell in enumerate(row):
                        col_widths[i] = max(col_widths[i], len(str(cell)))

                # Header
                header_line = " | ".join(h.ljust(w) for h, w in zip(headers, col_widths))
                lines.append(header_line)
                lines.append("-" * len(header_line))

                # Rows
                for row in rows:
                    row_line = " | ".join(str(cell).ljust(w) for cell, w in zip(row, col_widths))
                    lines.append(row_line)
                lines.append("")

            elif part_type == 'code':
                lang = part.get('language', '')
                lines.append(f"```{lang}")
                lines.append(part['text'])
                lines.append("```")
                lines.append("")

            elif part_type == 'divider':
                lines.append("-" * 60)
                lines.append("")

            elif part_type == 'blank':
                lines.append("")

        return "\n".join(lines)

    def build_xml(self) -> str:
        """
        Build OneNote XML content.

        Returns:
            str: XML content for OneNote
        """
        xml_parts = []

        for part in self.content_parts:
            part_type = part['type']

            if part_type == 'heading':
                text = self._escape_xml(part['text'])
                size = part.get('size', 16)
                xml_parts.append(
                    f'<one:OE><one:T style="font-size:{size}pt;font-weight:bold"><![CDATA[{text}]]></one:T></one:OE>'
                )

            elif part_type == 'text':
                text = self._escape_xml(part['text'])
                style = []
                if part.get('bold'):
                    style.append('font-weight:bold')
                if part.get('italic'):
                    style.append('font-style:italic')
                style_attr = f' style="{";".join(style)}"' if style else ''
                xml_parts.append(f'<one:OE><one:T{style_attr}><![CDATA[{text}]]></one:T></one:OE>')

            elif part_type in ['bullet_list', 'numbered_list']:
                for item in part['items']:
                    text = self._escape_xml(item)
                    xml_parts.append(f'<one:OE><one:T><![CDATA[• {text}]]></one:T></one:OE>')

            elif part_type == 'table':
                # OneNote tables are complex; for now, create simple text representation
                headers = part['headers']
                rows = part['rows']

                # Header row
                header_text = " | ".join(headers)
                xml_parts.append(f'<one:OE><one:T style="font-weight:bold"><![CDATA[{header_text}]]></one:T></one:OE>')

                # Data rows
                for row in rows:
                    row_text = " | ".join(str(cell) for cell in row)
                    xml_parts.append(f'<one:OE><one:T><![CDATA[{row_text}]]></one:T></one:OE>')

            elif part_type == 'code':
                code_text = self._escape_xml(part['text'])
                xml_parts.append(
                    f'<one:OE><one:T style="font-family:Courier New;font-size:10pt"><![CDATA[{code_text}]]></one:T></one:OE>'
                )

            elif part_type == 'divider':
                xml_parts.append('<one:OE><one:T><![CDATA[' + ('-' * 60) + ']]></one:T></one:OE>')

            elif part_type == 'blank':
                xml_parts.append('<one:OE><one:T><![CDATA[ ]]></one:T></one:OE>')

        return '\n'.join(xml_parts)

    @staticmethod
    def _escape_xml(text: str) -> str:
        """Escape XML special characters."""
        # CDATA handles most escaping, but still be safe
        return str(text).replace(']]>', ']]]]><![CDATA[>')

    def __str__(self) -> str:
        return self.build_simple()


class TemplateBuilder:
    """Pre-built templates for common OneNote pages."""

    @staticmethod
    def meeting_minutes(
        meeting_title: str,
        date: str,
        attendees: List[str],
        agenda_items: List[str],
        notes: str = "",
        action_items: List[Dict[str, str]] = None
    ) -> OneNoteContentBuilder:
        """
        Create meeting minutes template.

        Args:
            meeting_title: Meeting title
            date: Meeting date
            attendees: List of attendees
            agenda_items: Agenda items
            notes: Meeting notes
            action_items: List of action items with 'task', 'owner', 'due' keys

        Returns:
            OneNoteContentBuilder: Builder with meeting minutes content
        """
        builder = OneNoteContentBuilder(meeting_title)

        builder.add_text(f"Date: {date}", bold=True)
        builder.add_blank_line()

        builder.add_heading("Attendees", 2)
        builder.add_bullet_list(attendees)

        builder.add_heading("Agenda", 2)
        builder.add_numbered_list(agenda_items)

        if notes:
            builder.add_heading("Notes", 2)
            builder.add_text(notes)

        if action_items:
            builder.add_heading("Action Items", 2)
            headers = ["Task", "Owner", "Due Date"]
            rows = [[item.get('task', ''), item.get('owner', ''), item.get('due', '')] for item in action_items]
            builder.add_table(headers, rows)

        builder.add_divider()
        builder.add_text(f"Created: {datetime.now().strftime('%Y-%m-%d %H:%M')}", italic=True)

        return builder

    @staticmethod
    def project_status(
        project_name: str,
        status: str,
        progress: int,
        milestones: List[Dict[str, str]],
        issues: List[str] = None,
        next_steps: List[str] = None
    ) -> OneNoteContentBuilder:
        """
        Create project status template.

        Args:
            project_name: Project name
            status: Overall status (On Track, At Risk, Delayed)
            progress: Progress percentage (0-100)
            milestones: List of milestones with 'name', 'due', 'status' keys
            issues: List of issues
            next_steps: List of next steps

        Returns:
            OneNoteContentBuilder: Builder with project status content
        """
        builder = OneNoteContentBuilder(f"{project_name} - Status Report")

        builder.add_text(f"Status: {status}", bold=True)
        builder.add_text(f"Progress: {progress}%", bold=True)
        builder.add_text(f"Updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        builder.add_blank_line()

        builder.add_heading("Milestones", 2)
        headers = ["Milestone", "Due Date", "Status"]
        rows = [[m.get('name', ''), m.get('due', ''), m.get('status', '')] for m in milestones]
        builder.add_table(headers, rows)

        if issues:
            builder.add_heading("Issues & Risks", 2)
            builder.add_bullet_list(issues)

        if next_steps:
            builder.add_heading("Next Steps", 2)
            builder.add_numbered_list(next_steps)

        return builder

    @staticmethod
    def research_notes(
        topic: str,
        source: str,
        key_findings: List[str],
        details: str = "",
        tags: List[str] = None
    ) -> OneNoteContentBuilder:
        """
        Create research notes template.

        Args:
            topic: Research topic
            source: Information source
            key_findings: List of key findings
            details: Additional details
            tags: Tags for categorization

        Returns:
            OneNoteContentBuilder: Builder with research notes content
        """
        builder = OneNoteContentBuilder(f"Research: {topic}")

        builder.add_text(f"Source: {source}", bold=True)
        builder.add_text(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
        builder.add_blank_line()

        builder.add_heading("Key Findings", 2)
        builder.add_bullet_list(key_findings)

        if details:
            builder.add_heading("Details", 2)
            builder.add_text(details)

        if tags:
            builder.add_divider()
            builder.add_text(f"Tags: {', '.join(tags)}", italic=True)

        return builder

    @staticmethod
    def decision_log(
        decision_title: str,
        context: str,
        options: List[Dict[str, str]],
        chosen_option: str,
        rationale: str,
        stakeholders: List[str]
    ) -> OneNoteContentBuilder:
        """
        Create decision log template.

        Args:
            decision_title: Decision title
            context: Decision context
            options: List of options with 'name', 'pros', 'cons' keys
            chosen_option: Chosen option name
            rationale: Decision rationale
            stakeholders: List of stakeholders

        Returns:
            OneNoteContentBuilder: Builder with decision log content
        """
        builder = OneNoteContentBuilder(f"Decision: {decision_title}")

        builder.add_text(f"Date: {datetime.now().strftime('%Y-%m-%d')}", bold=True)
        builder.add_text(f"Status: Approved", bold=True)
        builder.add_blank_line()

        builder.add_heading("Context", 2)
        builder.add_text(context)

        builder.add_heading("Options Considered", 2)
        for option in options:
            builder.add_heading(option.get('name', ''), 3)
            builder.add_text(f"Pros: {option.get('pros', '')}")
            builder.add_text(f"Cons: {option.get('cons', '')}")
            builder.add_blank_line()

        builder.add_heading("Decision", 2)
        builder.add_text(f"Chosen: {chosen_option}", bold=True)
        builder.add_text(rationale)

        builder.add_heading("Stakeholders", 2)
        builder.add_bullet_list(stakeholders)

        return builder
