from pydantic import BaseModel, Field
from typing import List, Optional


from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_con import MsGraphInstance


class UserProfile(BaseModel):
    """MS Graph /me profile — maps camelCase JSON fields to snake_case attributes."""
    odata_context: str = Field(
        default='',
        alias='@odata.context',
        description='OData context URL'
    )
    business_phones: List[str] = Field(
        default_factory=list,
        alias='businessPhones',
        description='List of business phone numbers'
    )
    display_name: str = Field(
        default='',
        alias='displayName',
        description='Display name of the user'
    )
    given_name: Optional[str] = Field(
        default='',
        alias='givenName',
        description='Given (first) name of the user'
    )
    job_title: Optional[str] = Field(
        default=None,
        alias='jobTitle',
        description='Job title of the user'
    )
    mail: Optional[str] = Field(
        default=None,
        description='Primary email address of the user'
    )
    mobile_phone: Optional[str] = Field(
        default=None,
        alias='mobilePhone',
        description='Mobile phone number of the user'
    )
    office_location: Optional[str] = Field(
        default=None,
        alias='officeLocation',
        description='Office location of the user'
    )
    preferred_language: Optional[str] = Field(
        default=None,
        alias='preferredLanguage',
        description='Preferred language of the user'
    )
    surname: str = Field(
        default='',
        description='Surname (last name) of the user'
    )
    user_principal_name: str = Field(
        default='',
        alias='userPrincipalName',
        description='User principal name (UPN) of the user'
    )
    id: str = Field(
        default='',
        description='Unique identifier of the user'
    )

    class Config:
        populate_by_name = True


class ProfileHandler:

    def __init__(self, wui: "MsGraphInstance", me: Optional[UserProfile] = None):
        self.msg = wui
        self._me: Optional[UserProfile] = me

    @property
    def me(self) -> UserProfile | None:
        """Return the cached profile or None. Use me_async() to fetch from API."""
        return self._me

    async def me_async(self) -> UserProfile:
        if self._me is None:
            access_token = await self.msg.get_access_token_async()
            if not access_token:
                return UserProfile()
            response = await self.msg.run_async(url=f"{self.msg.msg_endpoint}me", token=access_token)
            if response is None or response.status_code != 200:
                return UserProfile()
            user_profile = response.json()
            profile =  UserProfile.model_validate(user_profile)
            self._me = profile
        return self._me
