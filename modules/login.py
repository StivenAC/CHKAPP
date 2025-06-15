import os
from ldap3 import Server, Connection, ALL, NTLM, SUBTREE
import logging
import json

def get_user_access(user_groups):
    """
    Determine the user's access configuration based on group memberships.
    """
    try:
        access_config = json.loads(os.getenv("ACCESS_CONFIG"))
    except json.JSONDecodeError:
        logging.error("Invalid JSON format in ACCESS_CONFIG environment variable.")
        return None

    # Match user groups with those in ACCESS_CONFIG
    for group, config in access_config.items():
        if any(group in user_group for user_group in user_groups):
            logging.info(f"Access granted based on group: {group}")
            return config

    logging.warning("No matching group found in ACCESS_CONFIG.")
    return None


def authenticate_user(username, password):
    server = Server(os.getenv("AD_SERVER"), get_info=ALL)
    user = f'{os.getenv("AD_DOMAIN")}\\{username}'

    try:
        # Attempt LDAP bind with provided credentials
        conn = Connection(server, user=user, password=password, authentication="NTLM", auto_bind=True)
        logging.info(f"LDAP bind successful for {username}.")

        # Search base for LDAP
        search_base = f"DC={os.getenv('AD_DOMAIN').replace('.', ',DC=')}"

        # Check if user is in allowed users
        allowed_users = os.getenv("ALLOWED_USERS").split(",")
        if username in allowed_users:
            return True

        # Search for the user's DN and group memberships
        conn.search(search_base, f"(sAMAccountName={username})", attributes=["distinguishedName", "memberOf"])

        if not conn.entries:
            logging.warning(f"User {username} not found in LDAP search.")
            return False

        # Retrieve user's group memberships
        user_groups = conn.entries[0].memberOf.values if conn.entries[0].memberOf else []

        # Determine access based on group memberships
        user_access = get_user_access(user_groups)
        if user_access:
            return user_access

        logging.warning(f"Access denied for {username}. No matching group configuration.")
        return False

    except Exception as e:
        logging.error(f"LDAP error for {username}: {e}")
        return False
