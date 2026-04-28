"""outlook_list_contacts, outlook_search_contacts, outlook_get_contact."""

from __future__ import annotations

from dataclasses import asdict
from typing import Any

from ..backends.base import Backend
from ..models.inputs import (
    GetContactInput,
    ListContactsInput,
    SearchContactsInput,
)


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_list_contacts",
        description="Paginated list of contacts in the default contacts folder.",
    )
    async def outlook_list_contacts(
        limit: int = 100, offset: int = 0
    ) -> dict[str, Any]:
        params = ListContactsInput(limit=limit, offset=offset)
        contacts = await backend.list_contacts(params.limit, params.offset)
        return {"contacts": [asdict(c) for c in contacts]}

    @mcp.tool(
        name="outlook_search_contacts",
        description="Keyword search across contact name, primary email, and company.",
    )
    async def outlook_search_contacts(
        query: str, limit: int = 50
    ) -> dict[str, Any]:
        params = SearchContactsInput(query=query, limit=limit)
        contacts = await backend.search_contacts(params.query, params.limit)
        return {"contacts": [asdict(c) for c in contacts]}

    @mcp.tool(
        name="outlook_get_contact",
        description="Return a single contact by id, including all phone numbers.",
    )
    async def outlook_get_contact(contact_id: str) -> dict[str, Any]:
        params = GetContactInput(contact_id=contact_id)
        contact = await backend.get_contact(params.contact_id)
        return asdict(contact)
