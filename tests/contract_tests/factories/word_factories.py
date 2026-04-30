"""
Word Data Factory

生成 Word 事件测试数据的工厂类。
"""

from __future__ import annotations

from typing import Any


class WordDataFactory:
    """
    Word 事件测试数据工厂。

    生成符合 Word Socket.IO 协议的测试数据。

    Examples:
        ```python
        factory = WordDataFactory()

        # 生成选中内容响应
        response = factory.selected_content_response(text="Hello World")

        # 生成插入文本响应
        response = factory.insert_text_response(success=True)
        ```
    """

    def selected_content_response(
        self,
        text: str | None = None,
        include_elements: bool = True,
        include_images: bool = False,
        include_tables: bool = False,
        character_count: int | None = None,
    ) -> dict[str, Any]:
        """
        生成 word:get:selectedContent 响应数据。

        Args:
            text: 选中的文本内容（默认使用测试文本）
            include_elements: 是否包含元素列表
            include_images: 是否包含图片
            include_tables: 是否包含表格
            character_count: 字符数（默认使用 text 的长度）

        Returns:
            符合协议的响应数据

        Examples:
            ```python
            factory = WordDataFactory()

            # 简单响应
            response = factory.selected_content_response(text="Hello World")

            # 完整响应（包含元素、图片、表格）
            response = factory.selected_content_response(
                text="Complete content",
                include_elements=True,
                include_images=True,
                include_tables=True
            )
            ```
        """
        if text is None:
            text = "Test selected content"

        if character_count is None:
            character_count = len(text)

        response: dict[str, Any] = {
            "text": text,
            "metadata": {
                "characterCount": character_count,
                "paragraphCount": 1,
                "tableCount": 1 if include_tables else 0,
                "imageCount": 1 if include_images else 0,
            },
        }

        # 添加元素列表
        if include_elements:
            elements: list[dict[str, Any]] = [
                {
                    "type": "paragraph",
                    "content": text,
                    "alignment": "left",
                    "style": "Normal",
                }
            ]

            # 添加图片元素
            if include_images:
                elements.append(
                    {
                        "type": "image",
                        "src": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
                        "width": 100,
                        "height": 100,
                    }
                )

            # 添加表格元素
            if include_tables:
                elements.append(
                    {
                        "type": "table",
                        "rows": 2,
                        "columns": 2,
                        "data": [["Cell 1", "Cell 2"], ["Cell 3", "Cell 4"]],
                    }
                )

            response["elements"] = elements
        else:
            response["elements"] = []

        return response

    def visible_content_response(
        self,
        text: str | None = None,
        include_elements: bool = True,
        include_images: bool = False,
        include_tables: bool = False,
        character_count: int | None = None,
        is_empty: bool = False,
    ) -> dict[str, Any]:
        """
        生成 word:get:visibleContent 响应数据。

        Args:
            text: 可见区域的文本内容（默认使用测试文本）
            include_elements: 是否包含元素列表
            include_images: 是否包含图片
            include_tables: 是否包含表格
            character_count: 字符数（默认使用 text 的长度）
            is_empty: 内容是否为空

        Returns:
            符合协议的响应数据

        Examples:
            ```python
            factory = WordDataFactory()

            # 简单响应
            response = factory.visible_content_response(text="Hello World")

            # 完整响应（包含元素、图片、表格）
            response = factory.visible_content_response(
                text="Complete visible content",
                include_elements=True,
                include_images=True,
                include_tables=True
            )

            # 空内容响应
            response = factory.visible_content_response(text="", is_empty=True)
            ```
        """
        if text is None:
            text = "Test visible content"

        if character_count is None:
            character_count = len(text)

        response: dict[str, Any] = {
            "text": text,
            "elements": [],
            "metadata": {
                "isEmpty": is_empty,
                "characterCount": character_count,
            },
        }

        # 添加元素列表
        if include_elements:
            elements: list[dict[str, Any]] = []

            # 添加文本元素
            if text:
                elements.append(
                    {
                        "type": "text",
                        "content": {"text": text},
                    }
                )

            # 添加图片元素
            if include_images:
                elements.append(
                    {
                        "type": "image",
                        "content": {
                            "base64": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
                            "width": 100,
                            "height": 100,
                        },
                    }
                )

            # 添加表格元素
            if include_tables:
                elements.append(
                    {
                        "type": "table",
                        "content": {
                            "rows": 2,
                            "columns": 2,
                            "data": [["Cell 1", "Cell 2"], ["Cell 3", "Cell 4"]],
                        },
                    }
                )

            response["elements"] = elements

        return response

    def selected_content_request(
        self,
        request_id: str = "test_req_001",
        document_uri: str = "file:///tmp/test.docx",
        options: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """
        生成 word:get:selectedContent 请求数据。

        Args:
            request_id: 请求 ID
            document_uri: 文档 URI
            options: 内容获取选项

        Returns:
            符合协议的请求数据
        """
        if options is None:
            options = {"includeText": True, "includeImages": False, "includeTables": False}

        return {
            "requestId": request_id,
            "documentUri": document_uri,
            "options": options,
        }

    def insert_text_response(
        self,
        success: bool = True,
        inserted_length: int | None = None,
        position: dict[str, int] | None = None,
    ) -> dict[str, Any]:
        """
        生成 word:insert:text 响应数据。

        Args:
            success: 是否成功
            inserted_length: 插入的文本长度
            position: 插入位置 {start, end}

        Returns:
            符合协议的响应数据
        """
        if not success:
            return {}

        if inserted_length is None:
            inserted_length = 11

        if position is None:
            position = {"start": 0, "end": 11}

        return {
            "insertedLength": inserted_length,
            "position": position,
        }

    def replace_selection_response(
        self,
        success: bool = True,
        replaced_length: int | None = None,
    ) -> dict[str, Any]:
        """
        生成 word:replace:selection 响应数据。

        Args:
            success: 是否成功
            replaced_length: 替换的文本长度

        Returns:
            符合协议的响应数据
        """
        if not success:
            return {}

        if replaced_length is None:
            replaced_length = 11

        return {
            "replacedLength": replaced_length,
        }

    def replace_text_response(
        self,
        replace_count: int = 5,
    ) -> dict[str, Any]:
        """
        生成 word:replace:text 响应数据。

        Args:
            replace_count: 替换的次数

        Returns:
            符合协议的响应数据

        Examples:
            ```python
            factory = WordDataFactory()

            # 默认响应（替换 5 次）
            response = factory.replace_text_response()

            # 自定义响应（替换 10 次）
            response = factory.replace_text_response(replace_count=10)

            # 无替换
            response = factory.replace_text_response(replace_count=0)
            ```
        """
        return {
            "replaceCount": replace_count,
        }

    def document_structure_response(
        self,
        paragraph_count: int = 10,
        table_count: int = 2,
        image_count: int = 1,
        section_count: int = 1,
    ) -> dict[str, Any]:
        """
        生成 word:get:documentStructure 响应数据。

        Args:
            paragraph_count: 段落数量
            table_count: 表格数量
            image_count: 图片数量
            section_count: 节数量

        Returns:
            符合协议的响应数据

        Examples:
            ```python
            factory = WordDataFactory()

            # 默认响应
            response = factory.document_structure_response()

            # 自定义响应
            response = factory.document_structure_response(
                paragraph_count=50,
                table_count=5,
                image_count=3,
                section_count=2
            )
            ```
        """
        return {
            "paragraphCount": paragraph_count,
            "tableCount": table_count,
            "imageCount": image_count,
            "sectionCount": section_count,
        }

    def document_stats_response(
        self,
        word_count: int = 1000,
        character_count: int = 5000,
        paragraph_count: int = 20,
    ) -> dict[str, Any]:
        """
        生成 word:get:documentStats 响应数据。

        Args:
            word_count: 字数
            character_count: 字符数
            paragraph_count: 段落数

        Returns:
            符合协议的响应数据

        Examples:
            ```python
            factory = WordDataFactory()

            # 默认响应
            response = factory.document_stats_response()

            # 自定义响应
            response = factory.document_stats_response(
                word_count=50000,
                character_count=250000,
                paragraph_count=1000
            )
            ```
        """
        return {
            "wordCount": word_count,
            "characterCount": character_count,
            "paragraphCount": paragraph_count,
        }

    def selection_response(
        self,
        is_empty: bool = False,
        selection_type: str = "Normal",
        start: int | None = None,
        end: int | None = None,
        text: str | None = None,
    ) -> dict[str, Any]:
        """
        生成 word:get:selection 响应数据。

        Args:
            is_empty: 选区是否为空（光标点）
            selection_type: 选区类型 (NoSelection, InsertionPoint, Normal)
            start: 起始位置（字符偏移）
            end: 结束位置（字符偏移）
            text: 选区文本内容

        Returns:
            符合协议的响应数据

        Examples:
            ```python
            factory = WordDataFactory()

            # 正常选区（有高亮文本）
            response = factory.selection_response(
                is_empty=False,
                selection_type="Normal",
                start=100,
                end=150,
                text="Selected text"
            )

            # 光标位置（无高亮）
            response = factory.selection_response(
                is_empty=True,
                selection_type="InsertionPoint",
                start=100,
                end=100
            )

            # 无选区
            response = factory.selection_response(
                is_empty=True,
                selection_type="NoSelection"
            )
            ```
        """
        response: dict[str, Any] = {
            "isEmpty": is_empty,
            "type": selection_type,
        }

        # 添加可选字段
        if start is not None:
            response["start"] = start
        if end is not None:
            response["end"] = end
        if text is not None:
            response["text"] = text

        return response

    def select_text_response(
        self,
        success: bool = True,
        match_count: int = 1,
        selected_index: int = 1,
        selected_text: str = "Selected text",
        selection_info: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """
        生成 word:select:text 响应数据。

        Args:
            success: 是否成功选中
            match_count: 找到的匹配项数量
            selected_index: 选中的匹配项索引（1-based）
            selected_text: 选中的文本
            selection_info: 选区详细信息

        Returns:
            符合协议的响应数据

        Examples:
            ```python
            factory = WordDataFactory()

            # 默认响应（选中1个匹配项）
            response = factory.select_text_response()

            # 自定义响应（选中第2个匹配项，共3个）
            response = factory.select_text_response(
                match_count=3,
                selected_index=2,
                selected_text="Hello World"
            )

            # 未找到匹配
            response = factory.select_text_response(
                success=False,
                match_count=0
            )
            ```
        """
        if not success:
            return {
                "success": False,
                "error": {"code": "3000", "message": "Office API error"},
            }

        if selection_info is None:
            selection_info = {
                "type": "Normal",
                "start": 100,
                "end": 100 + len(selected_text),
                "text": selected_text,
                "isEmpty": False,
            }

        return {
            "success": success,
            "matchCount": match_count,
            "selectedIndex": selected_index,
            "selectedText": selected_text,
            "selectionInfo": selection_info,
        }

    def error_response(
        self,
        code: str = "3002",
        message: str = "Selection is empty",
    ) -> dict[str, Any]:
        """
        生成错误响应数据。

        Args:
            code: 错误代码
            message: 错误消息

        Returns:
            符合协议的错误响应数据
        """
        return {
            "success": False,
            "error": {
                "code": code,
                "message": message,
            },
        }

    # ------------------------------------------------------------------
    # Table Operation Responses (OASP /word Draft, v0.2.0)
    # ------------------------------------------------------------------

    def merge_cells_response(
        self,
        table_id: str = "table-0",
        row_count: int = 1,
        column_count: int = 5,
    ) -> dict[str, Any]:
        """生成 word:merge:cells 响应数据。

        Args:
            table_id: 目标表格 ID
            row_count: 合并区域行数
            column_count: 合并区域列数

        Returns:
            协议响应数据 {tableId, requestedRange: {rowCount, columnCount}}
        """
        return {
            "tableId": table_id,
            "requestedRange": {
                "rowCount": row_count,
                "columnCount": column_count,
            },
        }

    def update_table_cell_response(
        self,
        table_id: str = "table-0",
        cells_updated: int = 1,
        row_count: int = 5,
        column_count: int = 4,
    ) -> dict[str, Any]:
        """生成 word:update:tableCell 响应数据 (OASP v0.2.0)。"""
        return {
            "tableId": table_id,
            "cellsUpdated": cells_updated,
            "rowCount": row_count,
            "columnCount": column_count,
        }

    def update_table_row_column_response(
        self,
        table_id: str = "table-0",
        cells_updated: int = 0,
        row_count: int = 5,
        column_count: int = 4,
    ) -> dict[str, Any]:
        """生成 word:update:tableRowColumn 响应数据 (OASP v0.2.0)。"""
        return {
            "tableId": table_id,
            "cellsUpdated": cells_updated,
            "rowCount": row_count,
            "columnCount": column_count,
        }

    def update_table_format_response(
        self,
        table_id: str = "table-0",
        row_count: int = 5,
        column_count: int = 4,
    ) -> dict[str, Any]:
        """生成 word:update:tableFormat 响应数据 (OASP v0.2.0)。"""
        return {
            "tableId": table_id,
            "rowCount": row_count,
            "columnCount": column_count,
        }
