"""
Contract Tests for word:replace:text

测试 word:replace:text 事件的完整请求-响应流程。

Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30801921
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试成功替换文本的完整流程。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应
    3. 连接到 Workspace
    4. Workspace 发送 word:replace:text 命令
    5. 验证返回的响应数据符合协议
    """
    # Arrange: 创建并配置 Mock 客户端
    expected_count = 7
    search_text = "old text"
    replace_text = "new text"

    def replace_text_response(request: dict) -> dict:
        """动态响应工厂，验证请求数据并返回响应"""
        # 验证请求数据
        assert "requestId" in request
        assert "documentUri" in request
        assert request["searchText"] == search_text
        assert request["replaceText"] == replace_text
        assert "options" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.replace_text_response(replace_count=expected_count),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    # 创建客户端
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    # 在连接之前注册响应
    client.register_response("word:replace:text", replace_text_response)

    # 连接到服务器
    await client.connect()

    try:
        # Act: 执行动作
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": search_text,
                "replaceText": replace_text,
                "options": {"replaceAll": True},
            },
        )
        result = await workspace.execute(action)

        # Assert: 验证结果
        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["replaceCount"] == expected_count
        assert result.data["replaceCount"] == 7

        # 验证 Mock 客户端接收到了请求
        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:replace:text"
        assert event_data["searchText"] == search_text
        assert event_data["replaceText"] == replace_text
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_no_matches(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试没有找到匹配项的替换操作。

    ⚠️ **规范层反例守护，勿改成错误用例**（oasp#23 §SEARCH_NO_MATCH / issue #90）：
    ``replaceCount: 0`` 是**良定义的零元结果**——「替换了 0 处」是完整、真实、可消费的答案，
    故 MUST 返回 ``success: true``，**不得**返回 ``3012 SEARCH_NO_MATCH``。判据不是「搜索是否
    为变更的锚点」（本事件的搜索同样是锚点），而是「零元有没有对应的良定义结果」。
    报成错误会迫使调用方从错误处理路径消费正常结果。
    对照：``word:insert:comment`` 的 searchText 零匹配**没有**良定义零元结果，故 MUST 判 3012
    （见 ``test_insert_comment.py::test_insert_comment_search_no_match_3012``）。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应（replaceCount=0）
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证返回 0 次替换
    """

    # Arrange: 创建并配置 Mock 客户端
    def zero_matches_response(request: dict) -> dict:
        """零匹配响应工厂"""
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.replace_text_response(replace_count=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:replace:text", zero_matches_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "nonexistent text",
                "replaceText": "replacement",
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["replaceCount"] == 0
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_with_options(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试带各种选项的文本替换。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应并验证选项
    3. 连接到 Workspace
    4. Workspace 发送带选项的命令
    5. 验证选项正确传递到 Mock 客户端
    """
    # Arrange
    expected_options = {
        "matchCase": True,
        "matchWholeWord": True,
        "replaceAll": False,
    }

    def response_with_options_validation(request: dict) -> dict:
        """验证选项的响应工厂"""
        # 验证选项正确传递
        assert "options" in request
        actual_options = request.get("options", {})

        # 验证我们关心的字段
        for key, value in expected_options.items():
            assert actual_options.get(key) == value, f"Expected {key}={value}, got {actual_options.get(key)}"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.replace_text_response(replace_count=1),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:replace:text", response_with_options_validation)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "Test",
                "replaceText": "Exam",
                "options": expected_options,
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["replaceCount"] == 1
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_with_format(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试带格式信息的文本替换（为已有文本添加格式）。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应并验证 format 字段正确传递
    3. 连接到 Workspace
    4. Workspace 发送带 format 的 word:replace:text 命令
    5. 验证 format 数据正确传递到 Mock 客户端
    """
    # Arrange (OASP 0.4.0: 字体收敛到 format.font = WordFont)
    expected_format = {
        "font": {"bold": True, "color": "#FF0000", "size": 16, "underline": "Single"},
    }

    def response_with_format_validation(request: dict) -> dict:
        """验证 format 字段的响应工厂"""
        # 验证 format 正确传递
        assert "format" in request, "format field should be present in request"
        actual_font = request["format"]["font"]
        assert actual_font["bold"] is True
        assert actual_font["color"] == "#FF0000"
        assert actual_font["size"] == 16
        assert actual_font["underline"] == "Single"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.replace_text_response(replace_count=3),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:replace:text", response_with_format_validation)
    await client.connect()

    try:
        # Act: 使用相同文本 + format 实现"为已有文本添加格式"
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "important",
                "replaceText": "important",
                "format": expected_format,
                "options": {"replaceAll": True},
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["replaceCount"] == 3

        # 验证 Mock 客户端接收到了带 format 的请求
        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:replace:text"
        assert event_data["searchText"] == "important"
        assert event_data["replaceText"] == "important"
        assert event_data["format"]["font"]["bold"] is True
        assert event_data["format"]["font"]["color"] == "#FF0000"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_with_style_name_format(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试使用 styleName 格式化文本。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应并验证 styleName 正确传递
    3. 连接到 Workspace
    4. 验证 styleName 数据正确传递
    """

    # Arrange
    def response_with_style_validation(request: dict) -> dict:
        """验证 styleName 的响应工厂"""
        assert "format" in request
        assert request["format"]["styleName"] == "Heading 1"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.replace_text_response(replace_count=1),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:replace:text", response_with_style_validation)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "Chapter Title",
                "replaceText": "Chapter Title",
                "format": {"styleName": "Heading 1"},
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["replaceCount"] == 1
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_without_format(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试不带 format 字段的替换（向后兼容）。

    验证 format 字段为可选，不提供时不影响原有功能。
    """

    # Arrange
    def response_without_format(request: dict) -> dict:
        """验证无 format 字段的响应工厂"""
        # format 应该不存在或为 None
        fmt = request.get("format")
        assert fmt is None, f"Expected format to be None, got {fmt}"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.replace_text_response(replace_count=2),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:replace:text", response_without_format)
    await client.connect()

    try:
        # Act: 不提供 format
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "old",
                "replaceText": "new",
                "options": {"replaceAll": True},
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["replaceCount"] == 2
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_missing_param_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """
    测试缺少必需参数的错误处理。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册错误响应（缺少参数）
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证错误信息正确返回
    """

    # Arrange: 创建并配置 Mock 客户端
    def error_response(request: dict) -> dict:
        """错误响应工厂"""
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "4001",
                "message": "Missing required parameters",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:replace:text", error_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "",  # Empty - should trigger error
                "replaceText": "",  # Empty - should trigger error
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is False
        assert "4001" in str(result.error) or "missing" in str(result.error).lower()
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_document_not_found(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """
    测试文档未找到的错误处理。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册错误响应（文档未找到）
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证错误信息正确返回
    """

    # Arrange
    def error_response(request: dict) -> dict:
        """错误响应工厂"""
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3001",
                "message": "Document not found",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/nonexistent.docx",
    )

    client.register_response("word:replace:text", error_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "test",
                "replaceText": "replacement",
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is False
        assert "3001" in str(result.error) or "not found" in str(result.error).lower()
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_replace_text_large_count(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试大量替换的响应。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应（大量替换）
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证大量替换计数正确返回
    """

    # Arrange
    large_count = 5000

    def large_replace_response(request: dict) -> dict:
        """大量替换响应工厂"""
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.replace_text_response(replace_count=large_count),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:replace:text", large_replace_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="replace:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "the",
                "replaceText": "a",
                "options": {"replaceAll": True},
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["replaceCount"] == large_count
    finally:
        await client.disconnect()
