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
        self.wb = None
        self.wks = None
        self.graph = None
        self.project_opened = False
        self.column_units = {}  # 存储列名到单位的映射
        self.column_order_from_dataname = []  # DataName行定义的列顺序

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

    def load_config(self, config_path):
        """加载绘图配置文件"""
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

            return config

        except Exception as e:
            print(f"❌ 解析配置文件时出错：{e}")
            return None

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
        print(f"   文件总行数: {len(lines)}")

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
                    op.new(file=project_path)
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

    def export_to_origin(self, df, data_filename, project_path=None, project_exists=False):
        """将DataFrame导出到Origin工作簿"""
        try:
            op.set_show(True)

            if not self.open_or_create_project(project_path, project_exists):
                return False

            book_name = self.get_safe_filename(os.path.basename(data_filename))
            sheet_name = f"{book_name}_Sheet"

            print(f"📘 创建工作簿: '{book_name}'")
            print(f"   导入 {df.shape[1]} 列: {list(df.columns)}")

            # 创建新工作簿
            self.wb = op.new_book('w', book_name)
            self.wks = self.wb[0]
            self.wks.name = sheet_name

            # 导出数据到Origin
            self.wks.from_df(df)

            # 设置列宽
            print(f"🔄 设置列格式...")
            try:
                self.wks.set_col_width(width=15)
                print(f"   已设置所有列宽为15")
            except:
                try:
                    # 备选方法：逐个设置列宽
                    for i in range(self.wks.cols):
                        self.wks.col(i).width = 15
                    print(f"   已逐个设置列宽为15")
                except Exception as width_error:
                    print(f"⚠️  设置列宽失败: {width_error}")

            # 添加单位信息到列标题 - 修复这里：使用 col() 而不是 cols()
            if hasattr(self, 'column_units') and self.column_units:
                print(f"📝 添加单位信息...")
                success_count = 0

                for i, col_name in enumerate(df.columns):
                    if col_name in self.column_units:
                        unit = self.column_units[col_name]

                        try:
                            # 获取列对象 - 关键修复：使用 col() 方法
                            col_obj = self.wks._find_col(i)

                            # 方法1：尝试设置单位
                            try:
                                #col_obj.units = unit
                                col_obj.SetUnits(unit)
                                print(f"   ✓ 列 {i}: {col_name} [单位属性: {unit}]")
                                success_count += 1
                            except:
                                # 方法2：尝试设置注释
                                try:
                                    col_obj.comments = f"单位: {unit}"
                                    print(f"   ✓ 列 {i}: {col_name} [单位: {unit}]")
                                    success_count += 1
                                except:
                                    # 方法3：尝试设置长名称
                                    try:
                                        col_obj.lname = f"{col_name} ({unit})"
                                        print(f"   ✓ 列 {i}: {col_name} -> {col_name} ({unit})")
                                        success_count += 1
                                    except:
                                        print(f"   ⚠️ 列 {i}: {col_name} - 所有单位设置方法都失败")

                        except Exception as col_error:
                            print(f"   ❌ 列 {i}: {col_name} - 获取列对象失败: {col_error}")

                print(f"   成功为 {success_count}/{len(self.column_units)} 个列添加单位信息")
            else:
                print(f"ℹ️  未找到列单位信息，使用原始列名")

            # 验证数据导入成功
            print(f"✅ 数据已成功导入Origin工作簿")
            print(f"   工作簿名称: {self.wb.name}")
            print(f"   工作表名称: {self.wks.name}")
            print(f"   数据维度: {len(df)} 行 × {len(df.columns)} 列")

            # 打印调试信息
            print(f"🔍 调试信息:")
            print(f"   工作表列数: {self.wks.cols}")
            if self.wks.cols > 0:
                try:
                    test_col = self.wks.col(0)
                    print(f"   第0列对象类型: {type(test_col)}")
                    print(f"   第0列名称: {test_col.name}")
                except Exception as debug_e:
                    print(f"   获取第0列信息失败: {debug_e}")

            return True

        except Exception as e:
            print(f"❌ 导出到Origin时出错：{e}")
            import traceback
            traceback.print_exc()
            return False

    def plot_in_origin(self, x_col, y_col, output_path, save_project=True, project_path=None):
        """在Origin中绘制图形"""
        try:
            # 获取列索引 - 修复这里：使用 col() 而不是 cols()
            try:
                col_names = [self.wks.col(i).name for i in range(self.wks.cols)]
            except AttributeError:
                # 备选方法：如果 name 属性不可用
                col_names = []
                for i in range(self.wks.cols):
                    try:
                        col_names.append(self.wks.col(i).name)
                    except:
                        # 如果无法获取列名，使用默认名称
                        col_names.append(f"Col_{i + 1}")

            print(f"📊 工作表列名: {col_names}")

            if x_col not in col_names:
                print(f"❌ 数据中不存在X轴列 '{x_col}'")
                print(f"   可用列: {col_names}")
                return False

            if y_col not in col_names:
                print(f"❌ 数据中不存在Y轴列 '{y_col}'")
                print(f"   可用列: {col_names}")
                return False

            x_idx = col_names.index(x_col)
            y_idx = col_names.index(y_col)

            print(f"✅ 找到列位置: {x_col}[索引{x_idx}], {y_col}[索引{y_idx}]")

            # 生成Graph名称
            graph_name = f"{y_col}-{x_col}"
            print(f"📈 创建Graph: '{graph_name}'")

            # 检查Graph是否已存在
            try:
                existing_graphs = op.lt_graph()
                if existing_graphs and graph_name in existing_graphs:
                    import time
                    timestamp = time.strftime("%H%M%S")
                    graph_name = f"{graph_name}_{timestamp}"
            except:
                pass

            # 创建图形
            self.graph = op.new_graph(template='line')

            # 设置Graph名称
            try:
                self.graph.name = graph_name
            except:
                print(f"⚠️  无法设置Graph名称，使用默认名称")

            gl = self.graph[0]

            # 添加绘图
            try:
                plot = gl.add_plot(self.wks, coly=y_idx, colx=x_idx, type='line')
                print(f"✅ 成功添加绘图: {y_col} vs {x_col}")
            except Exception as plot_error:
                print(f"❌ 添加绘图时出错: {plot_error}")
                # 尝试备选方法
                try:
                    plot = gl.add_plot(self.wks, coly=y_idx, colx=x_idx)
                    print(f"✅ 使用备选方法添加绘图成功")
                except:
                    return False

            # 设置图形属性
            gl.rescale()
            gl.label('X').text = x_col
            gl.label('Y').text = y_col
            gl.title = f'{y_col} vs {x_col}'

            # 设置线条样式
            try:
                plot.color = '#FF6600'  # 橙色
                plot.width = 2
            except:
                print(f"⚠️  无法设置线条样式，使用默认样式")

            print(f"✅ Graph '{graph_name}' 创建成功")

            # 保存图形
            try:
                self.graph.save_fig(output_path)
                print(f"✅ 图形已保存为PNG: {output_path}")
            except Exception as save_error:
                print(f"❌ 保存图形时出错: {save_error}")
                return False

            # 保存工程
            if save_project and project_path:
                try:
                    op.save(project_path)
                    print(f"✅ 工程文件已保存: {project_path}")
                except Exception as save_project_error:
                    print(f"⚠️  保存工程文件时出错: {save_project_error}")
                    # 尝试另存为
                    try:
                        backup_path = output_path.replace('.png', '_backup.opju')
                        op.save(backup_path)
                        print(f"✅ 工程文件已另存为: {backup_path}")
                    except:
                        print(f"⚠️  无法保存工程文件，请在Origin中手动保存")

            return True

        except Exception as e:
            print(f"❌ 绘图时出错：{e}")
            import traceback
            traceback.print_exc()
            return False

    def save_project_as(self, file_path):
        """另存工程文件"""
        try:
            op.save(file_path)
            print(f"✅ 工程已另存为: {file_path}")
            return True
        except Exception as e:
            print(f"❌ 保存工程时出错：{e}")
            return False

    def close_origin(self):
        """关闭Origin连接"""
        try:
            pass
        except:
            pass


def main():
    """主函数"""
    print("=" * 60)
    print("🏢 智能Origin数据处理与绘图系统")
    print("=" * 60)

    processor = OriginDataProcessor()

    if len(sys.argv) > 1:
        config_file = sys.argv[1]
        print(f"📋 使用命令行指定的配置文件: {config_file}")
    else:
        config_file = "OriginBook/plog_config.json"
        print(f"📋 使用默认配置文件: {config_file}")

    processor._config_dir = os.path.dirname(os.path.abspath(config_file)) if os.path.exists(
        config_file) else os.getcwd()

    try:
        # 1. 加载配置
        print("\n📋 步骤1: 加载绘图配置")
        config = processor.load_config(config_file)
        if not config:
            return False

        if not os.path.exists(config['data_file']):
            print(f"❌ 数据文件不存在: {config['data_file']}")
            return False

        graph_expected_name = f"{config['Y']}-{config['X']}"
        print(f"📊 预期生成的Graph名称: '{graph_expected_name}'")

        # 2. 读取数据文件
        print(f"\n📊 步骤2: 读取数据文件")
        need_columns = config.get('NeedCol')

        if need_columns:
            print(f"   配置要求提取列: {need_columns}")

        df, columns, success = processor.read_data_file(config['data_file'], need_columns)
        if not success:
            return False

        # 3. 导出到Origin
        print(f"\n🔄 步骤3: 导出数据到Origin")

        project_path = config.get('project')
        project_exists = config.get('project_exists', False)

        export_success = processor.export_to_origin(
            df,
            data_filename=config['data_file'],
            project_path=project_path,
            project_exists=project_exists
        )

        if not export_success:
            return False

        # 4. 绘制图形
        print(f"\n📈 步骤4: 根据配置绘制图形")
        print(f"   将生成Graph: '{graph_expected_name}'")

        data_basename = os.path.splitext(os.path.basename(config['data_file']))[0]
        output_dir = config['output_dir']
        png_output_path = os.path.join(output_dir, f"{data_basename}_plot.png")

        plot_success = processor.plot_in_origin(
            x_col=config['X'],
            y_col=config['Y'],
            output_path=png_output_path,
            save_project=True,
            project_path=project_path if project_exists else None
        )

        # 5. 保存新工程
        if not project_exists and project_path and plot_success:
            if not any(project_path.endswith(ext) for ext in ['.opj', '.opju']):
                project_path = project_path + '.opju'

            os.makedirs(os.path.dirname(project_path) or '.', exist_ok=True)

            try:
                op.save(project_path)
                print(f"✅ 新工程文件已保存: {project_path}")
            except Exception as e:
                print(f"❌ 保存新工程文件时出错: {e}")

        # 6. 显示结果
        print(f"\n{'=' * 60}")
        if plot_success:
            print("🎉 处理完成！")
            print(f"   数据文件: {config['data_file']}")
            print(f"   配置: X={config['X']}, Y={config['Y']}")
            print(f"   工作簿: 基于 '{os.path.basename(config['data_file'])}' 命名")
            print(f"   Graph: '{graph_expected_name}'")
            print(f"   输出图形: {png_output_path}")

            if project_path:
                if project_exists:
                    print(f"   工程文件: {project_path} (已更新)")
                else:
                    print(f"   工程文件: {project_path} (已创建)")
        else:
            print("⚠️  数据处理完成，但绘图失败")
            print("   数据已保存到Origin工作簿，请手动绘图")

        print("\n💡 提示：")
        print("1. 支持复杂的双列名定义格式")
        print("2. 自动提取并应用列单位信息")
        print("3. 以DataName行的列顺序为准")

        return plot_success

    except Exception as e:
        print(f"\n❌ 程序执行出错：{e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        processor.close_origin()


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
