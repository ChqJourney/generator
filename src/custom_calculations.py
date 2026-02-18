"""
自定义计算函数模块
用于long_term_table的高级数据处理
"""

try:
    from src.calculator import CalculationRegistry
    from src.utils.logging_config import get_logger
except ImportError:
    from calculator import CalculationRegistry
    from utils.logging_config import get_logger

logger = get_logger(__name__)


@CalculationRegistry.register("long_term_data_treatment")
def long_term_data_treatment(
    maintenance_table,
    photometric_data_table,
    calculated_column=1,
    photometric_column=3,
    decimal_places_config=None,
):
    """
    高级long_term_table数据处理器

    将maintenance_table和photometric_data_table合并处理，
    支持灵活的小数位数配置

    Args:
        maintenance_table: 维护数据表 (List[List[Any]])
        photometric_data_table: 光度数据表 (List[List[Any]])
        calculated_column: 计算结果存放的列索引 (默认1)
        photometric_column: photometric_table中用于计算的列索引 (默认3)
        decimal_places_config: 小数位数配置字典，支持以下格式:
            - 简单格式: {"4": 1, "5": 2}  # 第4列1位小数，第5列2位小数
            - 条件格式: {"4": {"condition": ">= 100", "true": 0, "false": 1}}

    Returns:
        List[List[Any]]: 处理后的表格数据
    """
    if decimal_places_config is None:
        decimal_places_config = {"4": 1}

    # 处理字典格式的输入（从report_data中提取value字段）
    if isinstance(maintenance_table, dict) and "value" in maintenance_table:
        maintenance_table = maintenance_table["value"]
    if isinstance(photometric_data_table, dict) and "value" in photometric_data_table:
        photometric_data_table = photometric_data_table["value"]

    result_table = []

    for i, maintenance_row in enumerate(maintenance_table):
        new_row = list(maintenance_row) if maintenance_row else []

        # 在指定位置插入新列（不是在末尾追加）
        # 如果 calculated_column 超出当前列数，则追加到末尾
        # 否则在 calculated_column 位置插入新列，原有列向后移动
        if calculated_column < len(new_row):
            # 在指定位置插入空值占位符，后续会填入计算结果
            new_row.insert(calculated_column, "")
        else:
            # calculated_column 超出范围，追加空列直到达到目标位置
            while len(new_row) <= calculated_column:
                new_row.append("")

        # 执行跨表计算
        if i < len(photometric_data_table):
            try:
                # 假设maintenance_table的第4列是基准值
                maintenance_value = float(maintenance_row[0])
                photometric_value = float(photometric_data_table[i][photometric_column])

                if photometric_value != 0:
                    calculated_value = 100 * maintenance_value / photometric_value

                    # 应用小数位数配置
                    formatted_value = _apply_decimal_config(
                        calculated_value, str(calculated_column), decimal_places_config
                    )

                    new_row[calculated_column] = formatted_value

            except (ValueError, TypeError, IndexError):
                # 计算失败时保持原值
                pass

        result_table.append(new_row)
    logger.debug("✅ 自定义计算: long_term_data_treatment 完成")
    logger.debug(f"   处理行数: {len(result_table)}")
    logger.debug(f"result:{result_table}")
    return result_table


def _apply_decimal_config(value, column_key, config):
    """
    根据配置应用小数位数

    Args:
        value: 数值
        column_key: 列的字符串索引（如"4"）
        config: 小数位数配置字典

    Returns:
        str: 格式化后的字符串
    """
    if column_key not in config:
        # 默认1位小数
        return f"{value:.1f}"

    config_value = config[column_key]

    # 简单格式：直接是整数
    if isinstance(config_value, int):
        return f"{value:.{config_value}f}"

    # 条件格式：字典包含condition
    if isinstance(config_value, dict):
        condition = config_value.get("condition", "")
        true_decimal = config_value.get("true", 0)
        false_decimal = config_value.get("false", 1)

        # 解析条件（简单支持 >, <, >=, <=, == ）
        if ">=" in condition:
            threshold = float(condition.replace(">=", "").strip())
            decimal = true_decimal if value >= threshold else false_decimal
        elif ">" in condition:
            threshold = float(condition.replace(">", "").strip())
            decimal = true_decimal if value > threshold else false_decimal
        elif "<=" in condition:
            threshold = float(condition.replace("<=", "").strip())
            decimal = true_decimal if value <= threshold else false_decimal
        elif "<" in condition:
            threshold = float(condition.replace("<", "").strip())
            decimal = true_decimal if value < threshold else false_decimal
        elif "==" in condition:
            threshold = float(condition.replace("==", "").strip())
            decimal = true_decimal if value == threshold else false_decimal
        else:
            decimal = false_decimal

        return f"{value:.{decimal}f}"

    # 默认返回
    return f"{value:.1f}"


@CalculationRegistry.register("calculate_rated_energy_efficacy")
def calculate_rated_energy_efficacy(photometric_data_table, digits=1):
    """
    计算额定能效

    计算光度数据表中第4列平均值除以第2列平均值

    Args:
        photometric_data_table: 光度数据表 (二维数组)
        digits: 小数位数，默认为1

    Returns:
        str: 额定能效值
    """
    # 处理字典格式的输入
    if isinstance(photometric_data_table, dict) and "value" in photometric_data_table:
        photometric_data_table = photometric_data_table["value"]

    if not photometric_data_table or len(photometric_data_table) == 0:
        return "N/A"

    try:
        col2_values = []
        col4_values = []

        for row in photometric_data_table:
            if len(row) > 3:  # 确保有第4列
                try:
                    val4 = float(row[3])
                    col4_values.append(val4)
                except (ValueError, TypeError):
                    pass

            if len(row) > 1:  # 确保有第2列
                try:
                    val2 = float(row[1])
                    col2_values.append(val2)
                except (ValueError, TypeError):
                    pass

        if len(col2_values) == 0 or len(col4_values) == 0:
            return "N/A"

        avg_col2 = sum(col2_values) / len(col2_values)
        avg_col4 = sum(col4_values) / len(col4_values)

        if avg_col2 == 0:
            return "N/A"

        energy_efficacy = avg_col4 / avg_col2
        return f"{energy_efficacy:.{digits}f}"

    except (ValueError, TypeError, ZeroDivisionError):
        return "N/A"


@CalculationRegistry.register("calculate_rated_energy_class_rating")
def calculate_rated_energy_class_rating(
    photometric_data_table, pon, useful_luminous_flux, non_directional, mains
):
    """
    计算额定能源等级评级

    根据 EU Regulation 2019/2015 Annex II Table 1 和 Table 2 计算光源的能效等级。

    计算公式:
    η_TM = (Φ_use / P_on) × F_TM

    其中:
    - Φ_use: 有用光通量 (lm)，取 photometric_data_table 第4列的平均值
    - P_on: 开启模式功率消耗 (W)，取 photometric_data_table 第2列的平均值
    - F_TM: 根据光源类型和电源类型的因子 (Table 2)
      * NDLS + MLS: 1.000
      * NDLS + NMLS: 0.926
      * DLS + MLS: 1.176
      * DLS + NMLS: 1.089

    Args:
        photometric_data_table: 光度数据表 (二维数组)
        pon: 功率 (保留参数，实际从表格计算)
        useful_luminous_flux: 有用光通量 (保留参数，实际从表格计算)
        non_directional: 是否非定向光源 (NDLS/DLS)
        mains: 是否市电供电 (MLS/NMLS)

    Returns:
        str: 额定能源等级 (A, B, C, D, E, F, G)
    """
    # 处理字典格式的输入
    if isinstance(photometric_data_table, dict) and "value" in photometric_data_table:
        photometric_data_table = photometric_data_table["value"]

    if not photometric_data_table or len(photometric_data_table) == 0:
        return "N/A"

    try:
        # 计算 Φ_use (第4列平均值) 和 P_on (第2列平均值)
        col2_values = []
        col4_values = []

        for row in photometric_data_table:
            if len(row) > 3:  # 确保有第4列
                try:
                    val4 = float(row[3])
                    col4_values.append(val4)
                except (ValueError, TypeError):
                    pass

            if len(row) > 1:  # 确保有第2列
                try:
                    val2 = float(row[1])
                    col2_values.append(val2)
                except (ValueError, TypeError):
                    pass

        if len(col2_values) == 0 or len(col4_values) == 0:
            return "N/A"

        phi_use = sum(col4_values) / len(col4_values)  # 有用光通量
        p_on = sum(col2_values) / len(col2_values)  # 功率

        if p_on == 0:
            return "N/A"

        # 解析 non_directional 和 mains 参数
        is_non_directional = _parse_checkbox_value(non_directional)
        is_mains = _parse_checkbox_value(mains)

        # 根据 Table 2 确定 F_TM 因子
        if is_non_directional:
            # NDLS (Non-Directional Light Source)
            if is_mains:
                f_tm = 1.000  # NDLS + MLS
            else:
                f_tm = 0.926  # NDLS + NMLS
        else:
            # DLS (Directional Light Source)
            if is_mains:
                f_tm = 1.176  # DLS + MLS
            else:
                f_tm = 1.089  # DLS + NMLS

        # 计算 η_TM = (Φ_use / P_on) × F_TM
        eta_tm = (phi_use / p_on) * f_tm

        # 根据 Table 1 确定能效等级
        if eta_tm >= 210:
            return "A"
        elif eta_tm >= 185:
            return "B"
        elif eta_tm >= 160:
            return "C"
        elif eta_tm >= 135:
            return "D"
        elif eta_tm >= 110:
            return "E"
        elif eta_tm >= 85:
            return "F"
        else:
            return "G"

    except (ValueError, TypeError, ZeroDivisionError):
        return "N/A"


@CalculationRegistry.register("format_sample_size")
def format_sample_size(containing_product, light_sources, controlgear_model):
    """
    格式化样本大小

    Args:
        containing_product: 包含产品信息
        light_sources: 光源信息
        controlgear_model: 控制装置型号

    Returns:
        str: 格式化后的样本大小
    """
    # TODO: 实现计算逻辑
    if controlgear_model:
        return "10 pcs per model + 3 pcs controlgear"
    else:
        return "10 pcs per model"


@CalculationRegistry.register("calculate_directional_info")
def calculate_directional_info(non_directional):
    """
    计算定向信息

    Args:
        non_directional: 非定向信息

    Returns:
        str: 定向信息
    """
    if non_directional is None:
        return "unknown"
    if isinstance(non_directional, dict) and "value" in non_directional:
        non_directional_value = non_directional["value"]
        if non_directional_value.lower() == "true":
            return "non-directional"
        elif non_directional_value.lower() == "false":
            return "directional"
    else:
        return "unknown"


@CalculationRegistry.register("calculate_ponmax")
def calculate_ponmax(
    useful_luminous_flux, non_directional, cri, led_source, color_tuneable, mains
):
    """
    计算 Ponmax (最大允许功率)

    根据 EU Regulation 2019/2020 法规计算光源的最大允许功率:
    Ponmax = C × (L + Φuse/(F × η)) × R

    Args:
        useful_luminous_flux: 有用光通量 (Φuse, 单位: lm)
        non_directional: 是否非定向光源 (True/False/字符串/checkbox字典)
        cri: 显色指数 (CRI)
        led_source: 是否为 LED 光源 (True/False/字符串/checkbox字典)
        color_tuneable: 是否可调色 (影响修正因子 C)
        mains: 是否为市电电源 (影响修正因子 C)

    Returns:
        str: Ponmax 值 (单位: W), 保留两位小数
    """
    # 解析输入值
    try:
        flux = float(useful_luminous_flux) if useful_luminous_flux else 0
    except (ValueError, TypeError):
        return "N/A"

    try:
        cri_val = float(cri) if cri else 80  # 默认 CRI 80
    except (ValueError, TypeError):
        cri_val = 80

    # 解析 checkbox 或布尔值
    is_non_directional = _parse_checkbox_value(non_directional)
    is_led = _parse_checkbox_value(led_source)
    is_color_tuneable = _parse_checkbox_value(color_tuneable)
    is_mains = _parse_checkbox_value(mains)

    # LED 光源参数 (根据 EU 2019/2020 Table 1 和 Table 2)
    # 阈值光效 (Threshold efficacy)
    eta = 120  # lm/W (LED 光源)

    # 终端损耗因子 (End loss factor)
    L = 2.0  # W (LED 光源)

    # 修正因子 C (根据光源类型和方向性)
    # Table 2: Basic values for correction factor C
    if is_non_directional:
        C = 1.0  # 非定向光源 (NDLS)
        F = 1.00  # 光效因子 (非定向光源使用总光通量)
    else:
        C = 0.79  # 定向光源 (DLS)
        F = 0.85  # 光效因子 (定向光源使用锥形光通量)

    # 可调色光源的特殊处理
    if is_color_tuneable:
        # 可调色光源使用 NDLS 的基本 C 值
        C = 1.0

    # CRI 因子 R 计算
    # 0.65 for CRI ≤ 25
    # (CRI + 80) / 160 for CRI > 25, rounded to two decimals
    if cri_val <= 25:
        R = 0.65
    else:
        R = round((cri_val + 80) / 160, 2)

    # 计算 Ponmax
    # Ponmax = C × (L + Φuse/(F × η)) × R
    try:
        ponmax = C * (L + flux / (F * eta)) * R
        return f"{ponmax:.2f}"
    except (ZeroDivisionError, ValueError, TypeError):
        return "N/A"


def _parse_checkbox_value(value):
    """
    解析 checkbox 值或布尔值

    Args:
        value: 可能是布尔值、字符串或 checkbox 字典格式 {"type": "checkbox", "value": "true"}

    Returns:
        bool: 解析后的布尔值
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.lower() in ("true", "1", "yes", "on", "是")
    if isinstance(value, dict):
        # 处理 checkbox 字典格式
        if value.get("type") == "checkbox":
            checkbox_val = value.get("value", "false")
            return str(checkbox_val).lower() in ("true", "1", "yes", "on")
        # 处理普通字典，检查是否有 'value' 键
        if "value" in value:
            val = value["value"]
            if isinstance(val, str):
                return val.lower() in ("true", "1", "yes", "on")
            return bool(val)
    return bool(value)


@CalculationRegistry.register("calculate_required_maintenance_percentage")
def calculate_required_maintenance_percentage(l70):
    """
    计算所需最小光通维持率 (Lumen Maintenance Factor)

    根据 EU Regulation 2019/2020 Annex II Table 4 计算 LED/OLED 光源
    在 3000 小时耐久测试后所需的最小光通维持率。

    公式: XLMF,MIN % = 100 × e^((3000 × ln(0.7)) / L70)

    其中:
    - L70: 声明的 L70B50 寿命（小时），即 50% 的光源光输出降级到
           初始光通量 70% 以下的时间

    Args:
        l70: 声明的 L70B50 寿命（小时, 在配置中传入 extracted_data.l70_lifetime

    Returns:
        str: 最小光通维持率百分比，保留一位小数，如 "85.0%"
    """
    import math

    # 解析 L70 寿命值
    try:
        if isinstance(l70, dict) and "value" in l70:
            l70 = float(l70["value"])
        else:
            l70 = float(l70)
    except (ValueError, TypeError):
        return "N/A"

    # 如果值太小（< 1000），可能是 rated_wattage 而非 L70
    # 此时使用默认 L70 值 25000 小时（标准 LED 寿命）
    if l70 < 1000:
        l70 = 25000  # 默认 L70B50 寿命（小时）

    # L70 必须为正数
    if l70 <= 0:
        return "N/A"

    # 根据 EU 2019/2020 Table 4 计算最小光通维持率
    # XLMF,MIN % = 100 × e^((3000 × ln(0.7)) / L70)
    xlmf_min = 100 * math.exp((3000 * math.log(0.7)) / l70)

    # 上限限制: 不得超过 96.0%
    if xlmf_min > 96.0:
        xlmf_min = 96.0

    return f"{xlmf_min:.1f}%"


@CalculationRegistry.register("calculate_light_source_tech")
def calculate_light_source_tech(led_source):
    """
    计算光源技术

    Args:
        led_source: LED 光源信息

    Returns:
        str: 光源技术类型
    """
    if led_source is None:
        return "unknown"
    if isinstance(led_source, dict) and "value" in led_source:
        led_source_value = led_source["value"]
        if led_source_value.lower() == "true":
            return "LED"
        elif led_source_value.lower() == "false":
            return "non-LED"
    else:
        return "unknown"


@CalculationRegistry.register("calculated_zone_table")
def calculated_zone_table(
    zone_0_30, zone_0_60, zone_0_90, zone_0_120, zone_0_180, beam_angel
):
    """
    计算区域表格数据

    Args:
        zone_0_30: 0-30度区域光通量
        zone_0_60: 0-60度区域光通量
        zone_0_90: 0-90度区域光通量
        zone_0_120: 0-120度区域光通量
        zone_0_180: 0-180度区域光通量
        beam_angel: 光束角

    Returns:
        List[List[Any]]: 区域表格数据
    """
    # find cover zone degree by beam angel, for example, if beam angel is 95, then cover zone degree is 0-120
    if beam_angel is not None:
        try:
            beam_angel_value = float(beam_angel)
            if beam_angel_value <= 30:
                return [["zone-0_30", zone_0_30]]
            elif beam_angel_value <= 60:
                return [["zone-0_30", zone_0_30], ["zone-0_60", zone_0_60]]
            elif beam_angel_value <= 90:
                return [
                    ["zone-0_30", zone_0_30],
                    ["zone-0_60", zone_0_60],
                    ["zone-0_90", zone_0_90],
                ]
            elif beam_angel_value <= 120:
                return [
                    ["zone-0_30", zone_0_30],
                    ["zone-0_60", zone_0_60],
                    ["zone-0_90", zone_0_90],
                    ["zone-0_120", zone_0_120],
                ]
            else:
                return [
                    ["zone-0_30", zone_0_30],
                    ["zone-0_60", zone_0_60],
                    ["zone-0_90", zone_0_90],
                    ["zone-0_120", zone_0_120],
                    ["zone-180", zone_0_180],
                ]
        except ValueError:
            pass


@CalculationRegistry.register("calculated_beam_table")
def calculated_beam_table(beam_angel, peak_intensity):
    """
    计算光束表格数据

    Args:
        beam_angel: 光束角度
        peak_intensity: 峰值强度

    Returns:
        List[List[Any]]: 光束表格数据
    """
    # TODO: 实现计算逻辑
    return [["beam angle", beam_angel], ["peak intensity", peak_intensity]]


@CalculationRegistry.register("calculate_tdb_remarks")
def calculate_tdb_remarks(is_init):
    """
    计算 TBD 备注信息

    Args:
        is_init: 是否为初始测试（布尔值或字符串/checkbox字典）

    Returns:
        str: 备注信息
    """
    is_initial = _parse_checkbox_value(is_init)
    if is_initial:
        return "except 3600hrs Lumen Maintenance test"
    else:
        return ""
