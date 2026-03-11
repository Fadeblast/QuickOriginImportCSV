# -*-coding =utf-8-*-
# @Time :2026/3/10 10:04
# @File :OriginBookProcessor.py
# @Software:PyCharm

import pandas as pd
import os
import json
import originpro as op
import sys
import re
import warnings

warnings.filterwarnings('ignore')


class OriginDataProcessor:
    """通用Origin数据处理与绘图类"""

    def __init__(self):
        self.workbooks = None
        self.wb = None
        self.wks = None
        self.graph = None
        self.project_opened = False
        self.column_units = {}  # 存储列名到单位的映射
        self.column_order_from_dataname = []  # DataName行定义的列顺序
        self._config_dir = None
        self.config = None

    def get_safe_filename(self, filename, max_length=30):
        """获取安全的工作簿名称"""
        base_name = os.path.splitext(filename)[0]
        safe_name = re.sub(r'[^\w\s\-]', '_', base_name)
        safe_name = re.sub(r'_+', '_', safe_name)

        if len(safe_name) > max_length:
            safe_name = safe_name[:max_length]

        if safe_name and safe_name[0].isdigit():
            safe_name = 'Data_' + safe_name

        if not safe_name or safe_name.isspace():
            safe_name = 'DataBook'

        return safe_name.strip()

    def is_origin_project_file(self, file_path):
        """检查文件是否是Origin工程文件"""
        if not file_path:
            return False

        valid_extensions = ['.opj', '.opju', '.ogg', '.ogw', '.otp', '.otpu']
        ext = os.path.splitext(file_path)[1].lower()
        return ext in valid_extensions

    def open_or_create_project(self, project_path, project_exists):
        """打开现有工程或创建新工程"""
        try:
            if project_path and project_exists:
                print(f"🔄 打开现有工程: {project_path}")
                op.open(file=project_path)
                self.project_opened = True
                print("✅ 工程已打开")
                return True
            else:
                if project_path:
                    print(f"📁 创建新工程: {project_path}")
                    project_dir = os.path.dirname(project_path)
                    if project_dir:
                        os.makedirs(project_dir, exist_ok=True)
                    op.new(asksave=project_path)
                    self.project_opened = True
                else:
                    print("📁 创建新工程")
                    self.project_opened = False

                return True

        except Exception as e:
            print(f"❌ 打开/创建工程时出错：{e}")
            import traceback
            traceback.print_exc()
            return False

    def find_origin_project(self, project_path):
        """智能查找Origin工程文件"""
        if not project_path:
            return None, False

        if os.path.exists(project_path):
            return project_path, True

        base_name = os.path.splitext(project_path)[0]
        possible_extensions = ['.opju', '.opj', '.ogg', '.ogw']

        for ext in possible_extensions:
            test_path = base_name + ext
            if os.path.exists(test_path):
                print(f"🔍 找到工程文件: {test_path}")
                return test_path, True

        dir_path = os.path.dirname(project_path) or '.'
        if os.path.exists(dir_path):
            for file in os.listdir(dir_path):
                file_base = os.path.splitext(file)[0]
                if file_base == os.path.basename(base_name):
                    full_path = os.path.join(dir_path, file)
                    if self.is_origin_project_file(full_path):
                        print(f"🔍 找到匹配的工程文件: {full_path}")
                        return full_path, True

        return project_path, False

    def _resolve_paths(self, config):
        """解析配置文件中的路径"""
        if hasattr(self, '_config_dir'):
            config_dir = self._config_dir
        else:
            config_dir = os.getcwd()

        if 'data_file' in config:
            data_path = config['data_file']
            if not os.path.isabs(data_path):
                data_path = os.path.join(config_dir, data_path)
            config['data_file'] = os.path.abspath(data_path)

        if 'output_dir' in config:
            output_path = config['output_dir']
            if not os.path.isabs(output_path):
                output_path = os.path.join(config_dir, output_path)
            config['output_dir'] = os.path.abspath(output_path)
            os.makedirs(config['output_dir'], exist_ok=True)

        return config

    def load_config(self, config_path):
        """加载配置文件"""
        if not os.path.exists(config_path):
            print(f"❌ 配置文件 '{config_path}' 不存在")
            return None

        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                content = f.read().strip()

            config = {}

            if content.startswith('{'):
                config = json.loads(content)
            else:
                for pair in content.split(','):
                    if ':' in pair:
                        key, value = pair.split(':', 1)
                        key = key.strip()
                        value = value.strip()

                        if key == 'NeedCol' and '|' in value:
                            config[key] = [col.strip() for col in value.split('|')]
                        else:
                            config[key] = value

            required_keys = ['data_file', 'X', 'Y']
            missing_keys = [key for key in required_keys if key not in config]

            if missing_keys:
                print(f"❌ 配置文件中缺少必需字段: {missing_keys}")
                print(f"必需字段: {required_keys}")
                return None

            if 'project' not in config:
                config['project'] = None

            if 'output_dir' not in config:
                data_dir = os.path.dirname(os.path.abspath(config['data_file']))
                config['output_dir'] = data_dir

            if 'NeedCol' in config:
                if isinstance(config['NeedCol'], str):
                    config['NeedCol'] = [col.strip() for col in config['NeedCol'].split(',')]
                elif not isinstance(config['NeedCol'], list):
                    print(f"⚠️  NeedCol格式错误，将提取全部列")
                    del config['NeedCol']

            if config['project']:
                actual_project_path, project_exists = self.find_origin_project(config['project'])
                config['project'] = actual_project_path
                config['project_exists'] = project_exists
            else:
                config['project_exists'] = False

            config = self._resolve_paths(config)

            print(f"✅ 成功加载配置：")
            print(f"   数据文件: {config['data_file']}")
            print(f"   X轴列: {config['X']}")
            print(f"   Y轴列: {config['Y']}")
            if 'NeedCol' in config:
                print(f"   需提取的列: {config['NeedCol']}")
            else:
                print(f"   需提取的列: 全部列")

            if config['project']:
                project_status = "已存在" if config['project_exists'] else "将创建"
                print(f"   工程文件: {config['project']} ({project_status})")
            else:
                print(f"   工程文件: 新建工程")

            print(f"   输出目录: {config['output_dir']}")
            self.config = config
            return config

        except Exception as e:
            print(f"❌ 解析配置文件时出错：{e}")
            return None

    def _detect_delimiter(self, line):
        """从行中检测分隔符"""
        if not line or not line.strip():
            return ','

        possible_delimiters = [',', '\t', ';', '|']
        delimiter_counts = {}

        for delim in possible_delimiters:
            count = line.count(delim)
            if count > 0:
                delimiter_counts[delim] = count

        if delimiter_counts:
            return max(delimiter_counts.items(), key=lambda x: x[1])[0]

        return ','

    def read_data_file(self, file_path, need_columns=None):
        """
        读取数据文件，处理复杂的双列名定义格式
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        except Exception as e:
            print(f"❌ 读取文件时发生错误：{e}")
            return None, None, False

        print(f"📂 开始解析文件: {os.path.basename(file_path)}")

        self.column_units.clear()
        self.column_order_from_dataname.clear()

        # 1. 首先找到AnalysisSetup...Datum.Name行提取单位映射
        datum_name_headers = []
        unit_mapping = {}

        for i, line in enumerate(lines):
            line = line.strip()

            if line.startswith('AnalysisSetup') and 'Datum.Name' in line:
                print(f"📋 找到AnalysisSetup...Datum.Name行 (第{i + 1}行)")

                # 检测分隔符
                delimiter = self._detect_delimiter(line)
                print(f"   检测到分隔符: '{delimiter}'")

                parts = line.split(delimiter)
                print(f"   分割结果: {parts}")

                if len(parts) > 2:
                    datum_name_headers = [h.strip() for h in parts[2:] if h.strip()]
                    print(f"   AnalysisSetup列名: {datum_name_headers}")

                    # 查找对应的单位行
                    for j in range(i + 1, min(i + 10, len(lines))):
                        unit_line = lines[j].strip()
                        if unit_line.startswith('AnalysisSetup') and 'Datum.Unit' in unit_line:
                            print(f"📏 找到对应的单位行 (第{j + 1}行)")
                            unit_parts = unit_line.split(delimiter)

                            if len(unit_parts) > 2:
                                units = [u.strip() for u in unit_parts[2:] if u.strip()]
                                print(f"   单位: {units}")

                                # 创建列名->单位的映射
                                for col, unit in zip(datum_name_headers, units):
                                    unit_mapping[col] = unit
                                    print(f"   {col} -> {unit}")
                            break
                break

        # 2. 找到DataName行（这是实际的列顺序）
        data_start_line = -1
        data_headers = []
        delimiter = ','  # 默认逗号分隔

        for i, line in enumerate(lines):
            line = line.strip()

            if line.startswith('DataName'):
                data_start_line = i
                print(f"\n📊 找到DataName行 (第{i + 1}行)")
                print(f"   行内容: {line}")

                # 检测DataName行的分隔符
                delimiter = self._detect_delimiter(line)
                print(f"   DataName行分隔符: '{delimiter}'")

                parts = line.split(delimiter)
                data_headers = [h.strip() for h in parts[1:] if h.strip()]
                self.column_order_from_dataname = data_headers.copy()
                print(f"   DataName列顺序: {data_headers}")
                print(f"   列数: {len(data_headers)}")
                break

        if data_start_line == -1:
            print("❌ 文件中未找到DataName行")
            return None, None, False

        # 3. 将单位映射应用到DataName列的顺序
        for col in data_headers:
            if col in unit_mapping:
                self.column_units[col] = unit_mapping[col]

        print(f"\n📏 最终列单位映射:")
        for col in data_headers:
            unit = self.column_units.get(col, '[无单位]')
            print(f"   {col}: {unit}")

        # 4. 确定需要提取的列
        if need_columns:
            valid_columns = []
            missing_columns = []

            for col in need_columns:
                if col in data_headers:
                    valid_columns.append(col)
                else:
                    missing_columns.append(col)

            if missing_columns:
                print(f"⚠️  以下需要的列在文件中不存在: {missing_columns}")
                print(f"   可用列: {data_headers}")
                print(f"   将提取文件中实际存在的列")

            if valid_columns:
                headers = valid_columns
                col_indices = [data_headers.index(col) for col in valid_columns]
                print(f"✅ 将提取 {len(valid_columns)} 列: {valid_columns}")
                print(f"   列索引: {col_indices}")
            else:
                print(f"⚠️  所有指定列都不存在，将提取所有列")
                headers = data_headers
                col_indices = list(range(len(data_headers)))
        else:
            headers = data_headers
            col_indices = list(range(len(data_headers)))
            print(f"✅ 将提取所有 {len(headers)} 列")

        # 5. 收集所有DataValue行数据
        data_rows = []
        skipped_rows = 0
        data_row_count = 0

        print(f"\n📥 开始读取DataValue数据...")

        for i in range(data_start_line + 1, len(lines)):
            line = lines[i].strip()

            if not line:
                skipped_rows += 1
                continue

            if line.startswith(('SetupTitle', 'PrimitiveTest', 'TestParameter',
                                'AnalysisSetup', 'Dimension1', 'Dimension2', '#')):
                skipped_rows += 1
                continue

            if not line.startswith('DataValue'):
                print(f"⚠️  第{i + 1}行不是DataValue行，跳过: {line[:50]}...")
                skipped_rows += 1
                continue

            # 解析DataValue行
            parts = line.split(delimiter)

            if len(parts) < 2:
                print(f"⚠️  第{i + 1}行DataValue格式错误，跳过")
                skipped_rows += 1
                continue

            # 跳过第一个"DataValue"，获取数据
            values = parts[1:]

            # 清理数据
            cleaned_values = []
            for val in values:
                cleaned_val = val.strip()
                # 处理科学计数法中的空格
                if ' ' in cleaned_val and ('E' in cleaned_val or 'e' in cleaned_val):
                    cleaned_val = cleaned_val.replace(' ', '')
                cleaned_values.append(cleaned_val)

            # 检查数据行是否足够长
            if len(cleaned_values) < len(data_headers):
                print(f"⚠️  第{i + 1}行数据列数不足 ({len(cleaned_values)} < {len(data_headers)})")
                # 用空值填充不足的部分
                cleaned_values.extend([''] * (len(data_headers) - len(cleaned_values)))

            # 提取指定列的数据
            selected_values = []
            for idx in col_indices:
                if idx < len(cleaned_values):
                    selected_values.append(cleaned_values[idx])
                else:
                    selected_values.append('')

            # 检查是否有实际数据
            has_data = any(val.strip() for val in selected_values)
            if has_data:
                data_rows.append(selected_values)
                data_row_count += 1

                if data_row_count <= 3:
                    print(f"   第{i + 1}行数据示例: {selected_values}")
            else:
                skipped_rows += 1

        if not data_rows:
            print("❌ 文件中未找到有效的DataValue数据")
            return None, None, False

        print(f"\n📈 数据读取统计:")
        print(f"   成功读取数据行: {len(data_rows)}")
        print(f"   跳过的行: {skipped_rows}")

        # 6. 创建DataFrame
        try:
            print(f"\n🔄 创建DataFrame...")
            df = pd.DataFrame(data_rows, columns=headers)
            print(f"   DataFrame创建成功，形状: {df.shape}")

            # 转换数据类型
            print(f"🔄 转换数据类型...")
            for col in headers:
                try:
                    # 尝试转换为数值类型
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    non_null = df[col].count()
                    null_count = len(df) - non_null
                    if null_count > 0:
                        print(f"   {col}: 数值类型, {null_count}个空值")
                    else:
                        print(f"   {col}: 数值类型")
                except:
                    print(f"   {col}: 保持为字符串类型")

            # 数据预览
            print(f"\n👀 数据预览 (前3行):")
            print(df.head(3).to_string())

            print(f"\n📊 数据摘要:")
            print(f"   总行数: {len(df)}")
            print(f"   总列数: {len(df.columns)}")
            print(f"   列名: {list(df.columns)}")

            numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
            if len(numeric_cols) > 0:
                print(f"   数值列统计:")
                for col in numeric_cols:
                    col_data = df[col].dropna()
                    if len(col_data) > 0:
                        print(f"     {col}: 范围=[{col_data.min():.4e}, {col_data.max():.4e}]")

            return df, headers, True

        except Exception as e:
            print(f"❌ 创建DataFrame时出错：{e}")
            import traceback
            traceback.print_exc()
            return None, None, False

    def get_csv_files(self, path):
        """返回指定路径中所有.csv文件的完整路径"""
        csv_files = []
        for root, dirs, files in os.walk(path):
            for file in files:
                if file.endswith('.csv'):
                    csv_files.append(os.path.join(root, file))
        return csv_files

    def export_to_origin(self, project_path=None, project_exists=False):
        """将DataFrame导出到Origin工作簿"""
        try:
            op.set_show(True)

            if not self.open_or_create_project(project_path, project_exists):
                return False

            csv_lists = self.get_csv_files(self.config['data_file'])
            print(csv_lists)
            need_columns = self.config.get('NeedCol')
            # book_name = self.get_safe_filename(os.path.basename(self.config.get('data_file')))
            book_name = 'OriginData'
            # 创建新工作簿
            self.wb = op.new_book('w', book_name)
            print(f"📘 创建工作簿: '{book_name}'")

            for index, data_item in enumerate(csv_lists, start=0):
                self.wb.add_sheet(active=False)  # 先增加工作表，虽然已经有一张表了，总保证总数多一张
                self.wks = self.wb[index]  # 使用一张

                df, columns, success = self.read_data_file(data_item, need_columns)
                if not success:
                    continue

                # 第一步：提取纯文件名（去掉路径），比如 "/path/to/example.csv" → "example.csv"
                basename = os.path.basename(data_item)
                # 第二步：拆分文件名和扩展名，比如 "example.csv" → ("example", ".csv")，取第一个元素
                filename_without_ext = os.path.splitext(basename)[0]
                sheet_name = f"{filename_without_ext}"  # 设置工作表名称
                self.wks.name = sheet_name

                print(f"   导入 {df.shape[1]} 列: {list(df.columns)}")

                # 存储所有工作簿的引用
                if not hasattr(self, 'workbooks'):
                    self.workbooks = []
                    self.workbooks.append(self.wb)

                # 导出数据到Origin
                self.wks.from_df(df)
                # 设置单位
                for i, col_name in enumerate(df.columns):
                    if col_name in self.column_units:
                        unit = self.column_units[col_name]
                        try:
                            # 获取列对象
                            col_obj = self.wks._find_col(i)
                            try:
                                col_obj.SetUnits(unit)
                                print(f"   ✓ 列 {i}: {col_name} [单位属性: {unit}]")
                            except:
                                print(f"   ⚠️ 列 {i}: {col_name} - 所有单位设置方法都失败")
                        except Exception as col_error:
                            print(f"   ❌ 列 {i}: {col_name} - 获取列对象失败: {col_error}")

                # 验证数据导入成功
                print(f"✅ 数据已成功导入Origin工作簿")
                print(f"   工作簿名称: {self.wb.name}")
                print(f"   工作表名称: {self.wks.name}")
                print(f"   数据维度: {len(df)} 行 × {len(df.columns)} 列")

                # 打印调试信息
                print(f"🔍 调试信息:")
                print(f"   工作表列数: {self.wks.cols}")

        except Exception as e:
            print(f"❌ 导出到Origin时出错：{e}")
            import traceback
            traceback.print_exc()
            return False
        self.opSave()

    def opSave(self):
        project_path = self.config.get('project')
        op.save(project_path)
        if op.oext:
            op.exit()


def main():
    processor = OriginDataProcessor()
    if len(sys.argv) > 1:
        config_file = sys.argv[1]
        print(f"📋 使用命令行指定的配置文件: {config_file}")
    else:
        config_file = "OriginBook/plog_config.json"
        print(f"📋 使用默认配置文件: {config_file}")

    processor._config_dir = os.path.dirname(os.path.abspath(config_file)) if os.path.exists(
        config_file) else os.getcwd()

    # 加载配置文件
    processor.load_config(config_file)

    # 处理数据
    project_path = processor.config.get('project')
    project_exists = processor.config.get('project_exists', False)
    ret = processor.export_to_origin(project_path, project_exists)

    return ret


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
