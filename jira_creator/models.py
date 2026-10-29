"""Only validated data can cross the extraction / Jira creation boundary."""

from typing import Literal
from pydantic import BaseModel, ConfigDict, Field, model_validator


class Model(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True, allow_inf_nan=False)


class CredentialsModel(Model):
    # Whitespace may be part of an actual password/API key.
    model_config = ConfigDict(str_strip_whitespace=False)


class Login(CredentialsModel):
    username: str = Field(min_length=1, max_length=80)
    password: str


class PasswordChange(CredentialsModel):
    current_password: str
    password: str


class UserCreate(CredentialsModel):
    username: str = Field(pattern=r"^[A-Za-z0-9_.@-]{2,80}$")
    display_name: str = Field(min_length=1, max_length=80)
    password: str
    role: Literal["admin", "user"] = "user"


class UserEdit(CredentialsModel):
    display_name: str = Field(min_length=1, max_length=80)
    role: Literal["admin", "user"]
    active: bool
    password: str | None = None


class Connection(CredentialsModel):
    name: str = Field(min_length=1, max_length=100)
    kind: Literal["jira", "llm"]
    base_url: str = Field(min_length=1, max_length=500)
    username: str = Field(default="", max_length=100)
    model: str = Field(default="", max_length=100)
    output_mode: Literal["json_object", "json_schema"] = "json_object"
    secret: str = Field(default="", max_length=4096)
    clear_secret: bool = False
    version: int | None = None


class FieldMapping(Model):
    project_key: str = Field(min_length=1, max_length=100)
    issuetype_id: str = Field(min_length=1, max_length=100)
    field_id: str = Field(pattern=r"^customfield_\d+$")
    label: str = Field(min_length=1, max_length=100)
    kind: Literal["text", "number", "date", "user", "users", "option", "options"]
    default: str | float | list[str] | None = None


class Defaults(Model):
    project_key: str = ""
    issuetype_id: str = ""
    priority_id: str = ""
    assignee: str | None = None
    due_date: str | None = None
    components: list[str] = Field(default_factory=list, max_length=100)
    labels: list[str] = Field(default_factory=list, max_length=100)


class Profile(Model):
    name: str = Field(min_length=1, max_length=100)
    jira_connection_id: str | None = None
    llm_connection_id: str | None = None
    defaults: Defaults = Field(default_factory=Defaults)
    members: dict[str, str] = Field(default_factory=dict, max_length=200)
    mappings: list[FieldMapping] = Field(default_factory=list, max_length=100)
    version: int | None = None


class Template(Model):
    name: str = Field(min_length=1, max_length=100)
    jira_base_url: str = Field(default="", max_length=500)
    llm_base_url: str = Field(default="", max_length=500)
    model: str = Field(default="", max_length=100)
    output_mode: Literal["json_object", "json_schema"] = "json_object"
    defaults: Defaults = Field(default_factory=Defaults)
    members: dict[str, str] = Field(default_factory=dict, max_length=200)
    mappings: list[FieldMapping] = Field(default_factory=list, max_length=100)
    version: int | None = None


class ProblemDetails(Model):
    category: Literal["frontend", "backend", "c_algorithm", "unclassified"] = "unclassified"
    category_reason: str = Field(default="", max_length=2000)
    business_module: str = Field(default="", max_length=200)
    reproduction_steps: list[str] = Field(default_factory=list, max_length=30)
    actual_result: str = Field(default="", max_length=5000)
    expected_result: str = Field(default="", max_length=5000)
    acceptance: str = Field(default="", max_length=5000)


class Issue(Defaults, ProblemDetails):
    item_id: str = Field(min_length=1, max_length=80, pattern=r"^[A-Za-z0-9_-]+$")
    summary: str = Field(min_length=1, max_length=255)
    description: str = Field(default="", max_length=30000)
    custom_fields: dict[str, str | float | list[str] | None] = Field(default_factory=dict, max_length=100)
    sources: list[str] = Field(default_factory=list, max_length=100)


class ExtractIssue(ProblemDetails):
    summary: str = Field(min_length=1, max_length=255)
    description: str = Field(default="", max_length=30000)
    assignee: str | None = None
    due_date: str | None = None
    sources: list[str] = Field(default_factory=list, max_length=100)


class Extraction(Model):
    issues: list[ExtractIssue] = Field(max_length=100)


class ExtractRequest(Model):
    profile_id: str
    raw_text: str = Field(min_length=1, max_length=100000)
    defaults: Defaults = Field(default_factory=Defaults)


class Batch(Model):
    profile_id: str
    defaults: Defaults = Field(default_factory=Defaults)
    issues: list[Issue] = Field(min_length=1, max_length=100)
    idempotency_key: str = Field(min_length=16, max_length=100, pattern=r"^[A-Za-z0-9_-]+$")


class DraftIssue(Issue):
    summary: str = Field(default="", max_length=255)


class Draft(Model):
    name: str = Field(min_length=1, max_length=100)
    profile_id: str | None = None
    raw_text: str = Field(default="", max_length=100000)
    defaults: Defaults = Field(default_factory=Defaults)
    # Drafts can be incomplete. The batch endpoint validates every Issue again.
    issues: list[DraftIssue] = Field(default_factory=list, max_length=100)
    version: int | None = None

    @model_validator(mode="after")
    def unique_items(self):
        ids = [issue.item_id for issue in self.issues]
        if len(ids) != len(set(ids)):
            raise ValueError("草稿中的事项标识不能重复")
        return self


class Resolve(Model):
    action: Literal["link", "retry"]
    issue_key: str = Field(default="", max_length=100, pattern=r"^([A-Za-z][A-Za-z0-9_]*-\d+)?$")
    confirmed_not_created: bool = False
    issue: Issue | None = None
