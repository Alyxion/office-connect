import asyncio
from typing import Callable, Dict, List
from office_con.web_user_instance import WebUserInstance

class BackgroundServiceRegistry:
    """Registry for background services that need to be notified of user events"""
    _instance = None

    def __init__(self):
        if BackgroundServiceRegistry._instance is not None:
            raise RuntimeError("Use BackgroundServiceRegistry.instance() to access singleton")
        
        self.login_callbacks: List[Callable[[WebUserInstance], None]] = []
        self.logout_callbacks: List[Callable[[str], None]] = []  # user_id
        self.refresh_callbacks: List[Callable[[WebUserInstance, str], None]] = []
        self.loop_callbacks: List[Callable[[WebUserInstance], None]] = []
        self.active_users: Dict[str, WebUserInstance] = {}

    @classmethod
    def instance(cls) -> 'BackgroundServiceRegistry':
        """Get singleton instance"""
        if cls._instance is None:
            cls._instance = BackgroundServiceRegistry()
        return cls._instance

    def on_login(self, callback: Callable[[WebUserInstance], None]):
        """Register callback for user login events"""
        self.login_callbacks.append(callback)

    def on_logout(self, callback: Callable[[str], None]):
        """Register callback for user logout events"""
        self.logout_callbacks.append(callback)

    async def notify_login(self, user: WebUserInstance):
        """Notify all registered callbacks about a user login"""
        if user.user_id is not None:
            self.active_users[user.user_id] = user
        for callback in self.login_callbacks:
            try:
                result = callback(user)
                if asyncio.iscoroutine(result):
                    await result
            except Exception as e:
                try:
                    print(f"Error in login callback: {e}")
                except Exception:
                    pass  # stdout may be broken (e.g. NiceGUI pipe)

    def notify_logout(self, user_id: str):
        """Notify all registered callbacks about a user logout"""
        if user_id in self.active_users:
            del self.active_users[user_id]
            for callback in self.logout_callbacks:
                try:
                    callback(user_id)
                except Exception as e:
                    print(f"Error in logout callback: {e}")

    async def notify_loop(self, user: WebUserInstance):
        """Notify all registered callbacks about a periodic loop update"""
        for callback in self.loop_callbacks:
            try:
                # check if its a coroutine
                result = callback(user)
                if asyncio.iscoroutine(result):
                    await result
            except Exception as e:
                print(f"Error in loop callback: {e}")

    def on_loop(self, callback: Callable[[WebUserInstance], None]):
        """Register callback for periodic loop updates"""
        self.loop_callbacks.append(callback)

    def get_active_users(self) -> Dict[str, WebUserInstance]:
        """Get all currently active users"""
        return self.active_users.copy()
