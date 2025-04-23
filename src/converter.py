import csv
import logging
import os
import docx
from . import utils  # 使用相对导入
from . import logger_config  # 使用相对导入

# 假设 Word 表头和 CSV 表头的常量已经定义
# 注意：这些常量应该与文档中的实际内容和期望的 CSV 输出一致
# 请根据实际情况修改这些常量
EXPECTED_WORD_HEADERS_NORMALIZED = [
    "序号",
    "资料名称",
    "资料类型",
    "保管部门",
    "保管期限",
    "保管责任人",
    "归档日期",
]
EXPECTED_CSV_HEADERS = [
    "DataID",
    "Name",
    "Type",
    "Department",
    "RetentionPeriod",
    "Custodian",
    "ArchiveDate",
]


class DocConverter:
    def __init__(self, word_path, csv_path):
        self.word_path = word_path
        self.csv_path = csv_path
        self.log_path = None  # 日志路径将在 setup_logger 中设置
        self.logger = None
        # self._setup_logger() # 在 convert 方法开始时调用以确保路径有效

    def _setup_logger(self):
        """配置日志记录器。"""
        try:
            # 从 CSV 文件名派生日志文件名
            log_dir = os.path.join(os.path.dirname(self.csv_path), "logs")
            csv_filename = os.path.splitext(os.path.basename(self.csv_path))[0]
            self.log_path = os.path.join(log_dir, f"{csv_filename}_conversion.log")
            self.logger = logger_config.setup_logging(self.log_path)
        except Exception as e:
            # 如果日志设置失败，提供一个基本的日志记录器以报告错误
            logging.basicConfig(
                level=logging.ERROR, format="%(asctime)s - %(levelname)s - %(message)s"
            )
            self.logger = logging.getLogger(__name__)  # 使用当前模块名
            self.logger.error(f"Failed to setup logger: {e}", exc_info=True)
            # 即使日志设置失败，也尝试继续，但记录会不完整

    def _check_word_table_header(self, table):
        """检查 Word 表格的表头是否符合预期。"""
        if not table.rows:
            if self.logger:
                self.logger.warning("Encountered a table with no rows.")
            return False
        try:
            header_cells = table.rows[0].cells
            actual_headers_normalized = [
                utils.normalize_header(cell.text) for cell in header_cells
            ]

            if len(actual_headers_normalized) != len(EXPECTED_WORD_HEADERS_NORMALIZED):
                if self.logger:
                    self.logger.warning(
                        f"Table header length mismatch. Expected: {len(EXPECTED_WORD_HEADERS_NORMALIZED)}, Found: {len(actual_headers_normalized)}. Headers: {actual_headers_normalized}"
                    )
                return False

            if actual_headers_normalized == EXPECTED_WORD_HEADERS_NORMALIZED:
                return True
            else:
                if self.logger:
                    self.logger.warning(
                        f"Table header content mismatch. Expected: {EXPECTED_WORD_HEADERS_NORMALIZED}, Found: {actual_headers_normalized}"
                    )
                return False
        except IndexError:
            if self.logger:
                self.logger.error(
                    "Error accessing table header cells. The table might be malformed.",
                    exc_info=True,
                )
            return False
        except Exception as e:
            if self.logger:
                self.logger.error(
                    f"Unexpected error checking Word table header: {e}", exc_info=True
                )
            return False

    def _check_csv_header(self):
        """检查 CSV 文件是否存在以及表头是否匹配。"""
        if not os.path.exists(self.csv_path):
            if self.logger:
                self.logger.info(
                    f"CSV file '{self.csv_path}' not found. Will create a new file."
                )
            return "create"

        try:
            with open(
                self.csv_path, "r", newline="", encoding="utf-8-sig"
            ) as f:  # utf-8-sig 读取带BOM的CSV
                reader = csv.reader(f)
                try:
                    header = next(reader)
                    actual_headers_normalized = [
                        utils.normalize_header(h) for h in header
                    ]

                    # CSV 表头检查：转换为小写进行比较可能更健壮
                    expected_headers_lower = [h.lower() for h in EXPECTED_CSV_HEADERS]
                    actual_headers_lower = [
                        h.lower() for h in actual_headers_normalized
                    ]

                    if actual_headers_lower == expected_headers_lower:
                        if self.logger:
                            self.logger.info(
                                f"CSV file '{self.csv_path}' exists with matching header. Will append data."
                            )
                        return "append"
                    else:
                        if self.logger:
                            self.logger.error(
                                f"CSV file '{self.csv_path}' header mismatch. Expected (case-insensitive): {expected_headers_lower}, Found: {actual_headers_lower}"
                            )
                        return "mismatch"
                except StopIteration:  # 文件是空的
                    if self.logger:
                        self.logger.warning(
                            f"CSV file '{self.csv_path}' exists but is empty. Will treat as new file."
                        )
                    return "create"
        except FileNotFoundError:
            if self.logger:
                self.logger.info(
                    f"CSV file '{self.csv_path}' not found during check. Will create a new file."
                )
            return "create"  # 万一在第一次检查和读取之间被删除
        except Exception as e:
            if self.logger:
                self.logger.error(
                    f"Error reading CSV file '{self.csv_path}': {e}", exc_info=True
                )
            return "error"

    def _extract_data_from_table(self, table, table_index):
        """从符合表头要求的表格中提取数据。"""
        extracted_rows = []
        original_row_indices = []  # 记录原始行号（基于1）
        try:
            # table.rows 包含表头行，索引从 0 开始
            for i, row in enumerate(table.rows):
                if i == 0:  # 跳过表头行
                    continue
                row_data = [cell.text for cell in row.cells]
                # 检查是否为空行（所有单元格都为空或仅包含空白字符）
                if all(not cell_text.strip() for cell_text in row_data):
                    if self.logger:
                        self.logger.debug(
                            f"Skipping empty row at table {table_index + 1}, original row index {i + 1}."
                        )
                    continue  # 跳过空行

                extracted_rows.append(row_data)
                original_row_indices.append(i + 1)  # Word 表格中原始行号（基于1）

            return extracted_rows, original_row_indices
        except Exception as e:
            if self.logger:
                self.logger.error(
                    f"Error extracting data from table {table_index + 1}: {e}",
                    exc_info=True,
                )
            return [], []  # 出错时返回空列表

    def _process_row(self, raw_row_data, table_index, row_index):
        """处理从 Word 提取的单行数据，转换为 CSV 格式。"""
        # 假设 EXPECTED_WORD_HEADERS_NORMALIZED 和 EXPECTED_CSV_HEADERS 长度一致
        # 并且 Word 列按顺序对应 CSV 列
        if len(raw_row_data) < len(EXPECTED_WORD_HEADERS_NORMALIZED):
            if self.logger:
                self.logger.warning(
                    f"Row {row_index} in table {table_index + 1} has fewer cells ({len(raw_row_data)}) than expected ({len(EXPECTED_WORD_HEADERS_NORMALIZED)}). Padding with empty strings. Data: {raw_row_data}"
                )
            # 用空字符串填充缺失的列
            raw_row_data.extend(
                [""] * (len(EXPECTED_WORD_HEADERS_NORMALIZED) - len(raw_row_data))
            )
        elif len(raw_row_data) > len(EXPECTED_WORD_HEADERS_NORMALIZED):
            if self.logger:
                self.logger.warning(
                    f"Row {row_index} in table {table_index + 1} has more cells ({len(raw_row_data)}) than expected ({len(EXPECTED_WORD_HEADERS_NORMALIZED)}). Truncating extra cells. Data: {raw_row_data}"
                )
            raw_row_data = raw_row_data[: len(EXPECTED_WORD_HEADERS_NORMALIZED)]

        processed_data = {}  # 使用字典更容易按 CSV 表头名称赋值

        # 按 Word 表头顺序提取数据
        try:
            data_name = raw_row_data[1].strip()  # 资料名称 (Word 第 2 列)
            data_type = raw_row_data[2].strip()  # 资料类型 (Word 第 3 列)
            department = raw_row_data[3].strip()  # 保管部门 (Word 第 4 列)
            retention_period = raw_row_data[4].strip()  # 保管期限 (Word 第 5 列)
            custodian = raw_row_data[5].strip()  # 保管责任人 (Word 第 6 列)
            archive_date_str = raw_row_data[6].strip()  # 归档日期 (Word 第 7 列)

            # 处理日期
            archive_date_obj = utils.parse_date(archive_date_str)
            if archive_date_obj:
                archive_date_formatted = archive_date_obj.strftime("%Y-%m-%d")
            else:
                # 如果日期解析失败，记录错误并认为此行处理失败
                if self.logger:
                    self.logger.error(
                        f"Failed to parse date '{archive_date_str}' in table {table_index + 1}, row {row_index}."
                    )
                return None  # 返回 None 表示处理失败

            # 填充字典，键使用 CSV 表头
            # DataID 通常需要生成或留空，这里暂时留空
            processed_data[EXPECTED_CSV_HEADERS[0]] = ""  # DataID
            processed_data[EXPECTED_CSV_HEADERS[1]] = data_name  # Name
            processed_data[EXPECTED_CSV_HEADERS[2]] = data_type  # Type
            processed_data[EXPECTED_CSV_HEADERS[3]] = department  # Department
            processed_data[EXPECTED_CSV_HEADERS[4]] = (
                retention_period  # RetentionPeriod
            )
            processed_data[EXPECTED_CSV_HEADERS[5]] = custodian  # Custodian
            processed_data[EXPECTED_CSV_HEADERS[6]] = (
                archive_date_formatted  # ArchiveDate
            )

            # 按照 CSV 表头顺序返回列表
            return [processed_data.get(header, "") for header in EXPECTED_CSV_HEADERS]

        except IndexError as e:
            if self.logger:
                self.logger.error(
                    f"Index error processing row {row_index} in table {table_index + 1}. Likely row structure issue. Data: {raw_row_data}. Error: {e}",
                    exc_info=True,
                )
            return None
        except Exception as e:
            if self.logger:
                self.logger.error(
                    f"Unexpected error processing row {row_index} in table {table_index + 1}. Data: {raw_row_data}. Error: {e}",
                    exc_info=True,
                )
            return None

    def convert(self):
        """执行 Word 到 CSV 的转换过程。"""
        self._setup_logger()  # 在开始时设置日志记录器
        if not self.logger:
            # 如果日志设置失败，无法继续，但应该尝试用 print 输出错误
            print("ERROR: Logger setup failed. Cannot proceed with detailed logging.")
            return {
                "status": "error",
                "message": "Logger setup failed. Cannot proceed.",
                "success": 0,
                "errors": 0,
                "skipped_empty": 0,
                "csv_path": self.csv_path,
                "log_path": self.log_path,
            }

        self.logger.info(
            f"Starting conversion from '{self.word_path}' to '{self.csv_path}'"
        )

        csv_mode = self._check_csv_header()

        if csv_mode == "mismatch" or csv_mode == "error":
            msg = f"CSV header check failed (mode: {csv_mode}). Please check the CSV file or logs."
            self.logger.error(msg)
            return {
                "status": "error",
                "message": msg,
                "success": 0,
                "errors": 0,
                "skipped_empty": 0,
                "csv_path": self.csv_path,
                "log_path": self.log_path,
            }

        success_count = 0
        error_count = 0
        skipped_empty_count = 0
        processed_data_for_csv = []  # 存储所有成功处理的行数据
        processed_tables = 0
        processed_rows_total = 0

        try:
            document = docx.Document(self.word_path)
            self.logger.info(f"Successfully opened Word document: '{self.word_path}'")
            self.logger.info(f"Found {len(document.tables)} tables in the document.")

            for table_index, table in enumerate(document.tables):
                self.logger.info(f"Processing table {table_index + 1}...")
                if self._check_word_table_header(table):
                    self.logger.info(
                        f"Table {table_index + 1} header matches. Extracting data..."
                    )
                    processed_tables += 1
                    extracted_rows, original_indices = self._extract_data_from_table(
                        table, table_index
                    )
                    self.logger.info(
                        f"Extracted {len(extracted_rows)} non-empty rows from table {table_index + 1}."
                    )

                    for i, raw_row in enumerate(extracted_rows):
                        original_row_index = original_indices[i]
                        processed_rows_total += 1
                        processed_row_data = self._process_row(
                            raw_row, table_index, original_row_index
                        )

                        if processed_row_data:
                            # 再次检查是否为空行 (理论上 _extract_data_from_table 已过滤, 但双重检查更安全)
                            if all(
                                not str(cell).strip() for cell in processed_row_data
                            ):
                                skipped_empty_count += 1
                                self.logger.info(
                                    f"Skipping empty row detected after processing: table {table_index + 1}, original row {original_row_index}."
                                )
                            else:
                                success_count += 1
                                processed_data_for_csv.append(processed_row_data)
                                self.logger.debug(
                                    f"Successfully processed row: table {table_index + 1}, original row {original_row_index}."
                                )
                        else:
                            error_count += 1
                            # 错误已在 _process_row 中记录

                else:
                    self.logger.warning(
                        f"Skipping table {table_index + 1} due to header mismatch."
                    )

        except docx.opc.exceptions.PackageNotFoundError:
            msg = f"Word document not found or invalid: '{self.word_path}'"
            self.logger.error(msg)
            return {
                "status": "error",
                "message": msg,
                "success": 0,
                "errors": error_count,
                "skipped_empty": skipped_empty_count,
                "csv_path": self.csv_path,
                "log_path": self.log_path,
            }
        except Exception as e:
            msg = f"An unexpected error occurred while reading the Word document: {e}"
            self.logger.error(msg, exc_info=True)
            return {
                "status": "error",
                "message": msg,
                "success": 0,
                "errors": error_count,
                "skipped_empty": skipped_empty_count,
                "csv_path": self.csv_path,
                "log_path": self.log_path,
            }

        # 检查是否有有效数据被处理
        if not processed_data_for_csv:
            if processed_tables == 0:
                msg = "No tables with matching headers found in the Word document."
                self.logger.warning(msg)
            elif error_count > 0:
                msg = f"Processed {processed_rows_total} rows from {processed_tables} matching tables, but all resulted in errors. Check logs for details."
                self.logger.warning(msg)
            elif skipped_empty_count > 0:
                msg = f"Processed {skipped_empty_count} rows from {processed_tables} matching tables, but all were empty or skipped. No data to write."
                self.logger.warning(msg)
            else:
                msg = "No processable data found in the Word document after checking tables and rows."
                self.logger.warning(msg)

            return {
                "status": "warning" if error_count == 0 else "error",
                "message": msg,
                "success": 0,
                "errors": error_count,
                "skipped_empty": skipped_empty_count,
                "csv_path": self.csv_path,
                "log_path": self.log_path,
            }

        # 写入 CSV 文件
        write_mode = "w" if csv_mode == "create" else "a"
        write_header = csv_mode == "create"
        try:
            # 确保目标目录存在
            os.makedirs(os.path.dirname(self.csv_path), exist_ok=True)

            with open(
                self.csv_path, write_mode, newline="", encoding="utf-8-sig"
            ) as f:  # 使用 utf-8-sig 写入带 BOM 的 CSV
                writer = csv.writer(f)
                if write_header:
                    writer.writerow(EXPECTED_CSV_HEADERS)
                    self.logger.info(
                        f"Writing header to new CSV file: {EXPECTED_CSV_HEADERS}"
                    )
                writer.writerows(processed_data_for_csv)
                self.logger.info(
                    f"Successfully wrote {success_count} rows to '{self.csv_path}'. Mode: {'w': 'overwrite/create', 'a': 'append'}[write_mode]."
                )

            final_message = f"Conversion finished. Success: {success_count}, Errors: {error_count}, Skipped Empty: {skipped_empty_count}."
            self.logger.info(final_message)
            return {
                "status": "success",
                "message": final_message,
                "success": success_count,
                "errors": error_count,
                "skipped_empty": skipped_empty_count,
                "csv_path": self.csv_path,
                "log_path": self.log_path,
            }

        except Exception as e:
            msg = f"Failed to write data to CSV file '{self.csv_path}': {e}"
            self.logger.error(msg, exc_info=True)
            return {
                "status": "error",
                "message": msg,
                "success": 0,
                "errors": error_count,
                "skipped_empty": skipped_empty_count,
                "csv_path": self.csv_path,
                "log_path": self.log_path,
            }


# 测试块 (可选，用于基本测试)
if __name__ == "__main__":
    # 创建虚拟 Word 和 CSV 文件进行测试
    print("Running basic converter test...")
    test_word_path = "test_converter_input.docx"
    test_csv_path = "test_converter_output.csv"
    test_log_dir = os.path.join(os.path.dirname(test_csv_path), "logs")
    test_log_path = os.path.join(
        test_log_dir,
        f"{os.path.splitext(os.path.basename(test_csv_path))[0]}_conversion.log",
    )

    # 准备 Word 内容
    doc = docx.Document()
    doc.add_paragraph("Test Document for Converter")
    # 表 1: 正确表头
    table1 = doc.add_table(rows=1, cols=len(EXPECTED_WORD_HEADERS_NORMALIZED))
    hdr_cells = table1.rows[0].cells
    for i, header_text in enumerate(EXPECTED_WORD_HEADERS_NORMALIZED):
        hdr_cells[i].text = header_text.upper()  # 测试大小写不敏感
    # 添加数据行
    row_cells = table1.add_row().cells
    row_cells[0].text = "1"
    row_cells[1].text = "Doc A"
    row_cells[2].text = "Type1"
    row_cells[3].text = "Dept X"
    row_cells[4].text = "10 Years"
    row_cells[5].text = "John Doe"
    row_cells[6].text = "2023-01-15"
    row_cells = table1.add_row().cells  # 空行
    row_cells = table1.add_row().cells
    row_cells[0].text = "2"
    row_cells[1].text = "Doc B"
    row_cells[2].text = "Type2"
    row_cells[3].text = "Dept Y"
    row_cells[4].text = "Permanent"
    row_cells[5].text = "Jane Smith"
    row_cells[6].text = "2024年2月"  # 测试年月格式
    row_cells = table1.add_row().cells
    row_cells[0].text = "3"
    row_cells[1].text = "Doc C"
    row_cells[2].text = "Type3"
    row_cells[3].text = "Dept Z"
    row_cells[4].text = "5 Years"
    row_cells[5].text = "Peter Pan"
    row_cells[6].text = "invalid-date"  # 测试无效日期

    # 表 2: 错误表头
    doc.add_paragraph("\n")
    table2 = doc.add_table(rows=1, cols=3)
    table2.rows[0].cells[0].text = "Wrong"
    table2.rows[0].cells[1].text = "Header"
    table2.rows[0].cells[2].text = "Format"
    table2.add_row().cells[0].text = "Data"

    try:
        doc.save(test_word_path)
        print(f"Created test Word file: {test_word_path}")

        # 删除旧的 CSV 和日志文件（如果存在）
        if os.path.exists(test_csv_path):
            os.remove(test_csv_path)
        if os.path.exists(test_log_path):
            os.remove(test_log_path)
        if not os.path.exists(test_log_dir):
            os.makedirs(test_log_dir)

        # 执行转换 (第一次, 创建模式)
        print("\n--- Running conversion (create mode) ---")
        converter = DocConverter(test_word_path, test_csv_path)
        result1 = converter.convert()
        print(f"Conversion Result 1: {result1}")

        # 再次执行转换 (追加模式)
        print("\n--- Running conversion (append mode) ---")
        # 修改 Word 文件以模拟新数据 (可选)
        # table1.add_row().cells[1].text = 'Doc D' ...
        # doc.save(test_word_path)
        converter2 = DocConverter(test_word_path, test_csv_path)
        result2 = converter2.convert()
        print(f"Conversion Result 2: {result2}")

        # 检查 CSV 内容
        if os.path.exists(test_csv_path):
            print(f"\n--- Content of {test_csv_path} ---")
            with open(test_csv_path, "r", encoding="utf-8-sig") as f:
                print(f.read())
        else:
            print(f"\n{test_csv_path} was not created.")

        # 检查日志内容
        if result1 and result1.get("log_path") and os.path.exists(result1["log_path"]):
            print(f"\n--- Content of {result1['log_path']} ---")
            with open(result1["log_path"], "r", encoding="utf-8") as f:
                print(f.read())
        elif result1:
            print(
                f"\nLog file path {result1.get('log_path')} invalid or file not created."
            )
        else:
            print("\nLog file path could not be determined from result1.")

    except Exception as e:
        print(f"An error occurred during testing: {e}")
    finally:
        # 清理测试文件 (可选)
        # if os.path.exists(test_word_path): os.remove(test_word_path)
        # if os.path.exists(test_csv_path): os.remove(test_csv_path)
        # if os.path.exists(test_log_path): os.remove(test_log_path)
        # if os.path.exists(test_log_dir): # 检查目录是否为空再删除
        #     try: os.rmdir(test_log_dir)
        #     except OSError: print(f"Could not remove log directory {test_log_dir} as it might not be empty.")
        # print("Cleaned up test files.")
        pass
