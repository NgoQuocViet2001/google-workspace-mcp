from __future__ import annotations

from typing import Any

from .common import compact_dict


def simplify_chat_user(user: dict[str, Any] | None) -> dict[str, Any]:
    if not user:
        return {}
    return compact_dict(
        {
            "name": user.get("name"),
            "display_name": user.get("displayName"),
            "type": user.get("type"),
            "email": user.get("email"),
            "domain_id": user.get("domainId"),
            "avatar_url": user.get("avatarUrl"),
            "is_anonymous": user.get("isAnonymous"),
        }
    )


def simplify_chat_attachment(attachment: dict[str, Any]) -> dict[str, Any]:
    drive_data = attachment.get("driveDataRef", {})
    return compact_dict(
        {
            "name": attachment.get("name"),
            "content_name": attachment.get("contentName"),
            "content_type": attachment.get("contentType"),
            "source": attachment.get("source"),
            "download_uri": attachment.get("downloadUri"),
            "thumbnail_uri": attachment.get("thumbnailUri"),
            "attachment_data_ref": attachment.get("attachmentDataRef"),
            "drive_data_ref": compact_dict(
                {
                    "drive_file_id": drive_data.get("driveFileId"),
                    "mime_type": drive_data.get("mimeType"),
                    "resource_type": drive_data.get("resourceType"),
                }
            )
            or None,
        }
    )


def simplify_emoji_reaction_summary(summary: dict[str, Any]) -> dict[str, Any]:
    return compact_dict(
        {
            "emoji": summary.get("emoji"),
            "reaction_count": summary.get("reactionCount"),
        }
    )


def simplify_chat_space(space: dict[str, Any]) -> dict[str, Any]:
    details = space.get("spaceDetails", {})
    return compact_dict(
        {
            "name": space.get("name"),
            "display_name": space.get("displayName"),
            "space_type": space.get("spaceType"),
            "single_user_bot_dm": space.get("singleUserBotDm"),
            "threading_state": space.get("spaceThreadingState"),
            "history_state": space.get("spaceHistoryState"),
            "external_user_allowed": space.get("externalUserAllowed"),
            "import_mode": space.get("importMode"),
            "admin_installed": space.get("adminInstalled"),
            "create_time": space.get("createTime"),
            "last_active_time": space.get("lastActiveTime"),
            "space_uri": space.get("spaceUri"),
            "customer": space.get("customer"),
            "description": details.get("description"),
            "guidelines": details.get("guidelines"),
            "access_settings": space.get("accessSettings"),
            "permission_settings": space.get("permissionSettings"),
            "predefined_permission_settings": space.get("predefinedPermissionSettings"),
        }
    )


def simplify_chat_membership(membership: dict[str, Any]) -> dict[str, Any]:
    group_member = membership.get("groupMember") or membership.get("group") or {}
    return compact_dict(
        {
            "name": membership.get("name"),
            "state": membership.get("state"),
            "role": membership.get("role"),
            "create_time": membership.get("createTime"),
            "delete_time": membership.get("deleteTime"),
            "member": simplify_chat_user(membership.get("member")) or None,
            "group_member": compact_dict(
                {
                    "name": group_member.get("name"),
                }
            )
            or None,
        }
    )


def simplify_chat_message(message: dict[str, Any]) -> dict[str, Any]:
    return compact_dict(
        {
            "name": message.get("name"),
            "sender": simplify_chat_user(message.get("sender")) or None,
            "create_time": message.get("createTime"),
            "last_update_time": message.get("lastUpdateTime"),
            "delete_time": message.get("deleteTime"),
            "text": message.get("text"),
            "formatted_text": message.get("formattedText"),
            "argument_text": message.get("argumentText"),
            "thread": compact_dict(
                {
                    "name": message.get("thread", {}).get("name"),
                    "thread_key": message.get("thread", {}).get("threadKey"),
                }
            )
            or None,
            "space": compact_dict({"name": message.get("space", {}).get("name")}) or None,
            "thread_reply": message.get("threadReply"),
            "client_assigned_message_id": message.get("clientAssignedMessageId"),
            "annotations": message.get("annotations"),
            "attachments": [simplify_chat_attachment(item) for item in message.get("attachment", [])] or None,
            "cards_v2": message.get("cardsV2"),
            "accessory_widgets": message.get("accessoryWidgets"),
            "emoji_reactions": [
                simplify_emoji_reaction_summary(item) for item in message.get("emojiReactionSummaries", [])
            ]
            or None,
            "private_message_viewer": simplify_chat_user(message.get("privateMessageViewer")) or None,
            "deletion_metadata": message.get("deletionMetadata"),
            "quoted_message_metadata": message.get("quotedMessageMetadata"),
            "matched_url": message.get("matchedUrl"),
            "attached_gifs": message.get("attachedGifs"),
        }
    )
