"""
单元测试：custom_calculations.py
为自定义计算函数提供全面的单元测试
"""

import pytest
import sys
import math
from unittest.mock import patch, MagicMock

# 导入被测试的模块
sys.path.insert(0, 'src')
from custom_calculations import (
    long_term_data_treatment,
    _apply_decimal_config,
    calculate_rated_energy_efficacy,
    calculate_rated_energy_class_rating,
    format_sample_size,
    calculate_directional_info,
    calculate_ponmax,
    _parse_checkbox_value,
    calculate_required_maintenance_percentage,
    calculate_light_source_tech,
    calculated_zone_table,
    calculated_beam_table,
    calculate_tdb_remarks,
)


# =============================================================================
# long_term_data_treatment 函数测试
# =============================================================================

class TestLongTermDataTreatment:
    """long_term_data_treatment 函数的单元测试"""
    
    def test_basic_calculation(self):
        """测试基本跨表计算"""
        maintenance_table = [[100.0], [200.0], [300.0]]
        photometric_data_table = [
            [0, 50.0, 0, 1000.0],
            [0, 60.0, 0, 2000.0],
            [0, 70.0, 0, 3000.0]
        ]
        
        result = long_term_data_treatment(
            maintenance_table, 
            photometric_data_table,
            calculated_column=1,
            photometric_column=3
        )
        
        # 100*100/1000 = 10.0, 100*200/2000 = 10.0, 100*300/3000 = 10.0
        assert len(result) == 3
        assert result[0][1] == "10.0"
        assert result[1][1] == "10.0"
        assert result[2][1] == "10.0"
    
    def test_with_dict_input(self):
        """测试字典格式的输入"""
        maintenance_table = {'value': [[100.0], [200.0]]}
        photometric_data_table = {'value': [[0, 50.0, 0, 1000.0], [0, 60.0, 0, 2000.0]]}
        
        result = long_term_data_treatment(maintenance_table, photometric_data_table)
        
        assert len(result) == 2
        assert result[0][1] == "10.0"
    
    def test_with_decimal_config(self):
        """测试带小数位数配置"""
        maintenance_table = [[100.0], [200.0]]
        photometric_data_table = [
            [0, 50.0, 0, 1000.0],
            [0, 60.0, 0, 2000.0]
        ]
        decimal_config = {"1": 2}
        
        result = long_term_data_treatment(
            maintenance_table,
            photometric_data_table,
            calculated_column=1,
            decimal_places_config=decimal_config
        )
        
        assert result[0][1] == "10.00"
        assert result[1][1] == "10.00"
    
    def test_calculated_column_beyond_range(self):
        """测试 calculated_column 超出当前列数的情况"""
        maintenance_table = [[100.0]]  # 只有1列
        photometric_data_table = [[0, 50.0, 0, 1000.0]]
        
        result = long_term_data_treatment(
            maintenance_table,
            photometric_data_table,
            calculated_column=5  # 超出范围
        )
        
        assert len(result[0]) == 6  # 应该有6列（索引0-5）
        assert result[0][5] == "10.0"  # 在索引5处插入结果
    
    def test_empty_maintenance_table(self):
        """测试空的维护表"""
        result = long_term_data_treatment([], [])
        assert result == []
    
    def test_photometric_table_shorter(self):
        """测试光度数据表比维护表短"""
        maintenance_table = [[100.0], [200.0], [300.0]]
        photometric_data_table = [[0, 50.0, 0, 1000.0]]  # 只有1行
        
        result = long_term_data_treatment(maintenance_table, photometric_data_table)
        
        assert len(result) == 3
        assert result[0][1] == "10.0"  # 第一行有计算
        assert result[1][1] == ""  # 第二行没有对应的光度数据，保持空值
    
    def test_zero_photometric_value(self):
        """测试光度值为零的情况"""
        maintenance_table = [[100.0]]
        photometric_data_table = [[0, 50.0, 0, 0.0]]  # 第4列为0
        
        result = long_term_data_treatment(maintenance_table, photometric_data_table)
        
        # 除以零，保持空值
        assert result[0][1] == ""
    
    def test_invalid_value_handling(self):
        """测试无效值处理"""
        maintenance_table = [["invalid"]]
        photometric_data_table = [[0, 50.0, 0, 1000.0]]
        
        result = long_term_data_treatment(maintenance_table, photometric_data_table)
        
        # 无效值导致计算失败，保持空值
        assert result[0][1] == ""
    
    def test_default_decimal_config(self):
        """测试默认小数位数配置"""
        maintenance_table = [[100.0]]
        photometric_data_table = [[0, 50.0, 0, 1000.0]]
        
        result = long_term_data_treatment(
            maintenance_table,
            photometric_data_table,
            decimal_places_config=None  # 使用默认
        )
        
        assert result[0][1] == "10.0"  # 默认1位小数


# =============================================================================
# _apply_decimal_config 函数测试
# =============================================================================

class TestApplyDecimalConfig:
    """""_apply_decimal_config 函数的单元测试"""
    
    def test_simple_int_config(self):
        """测试简单整数配置"""
        config = {"4": 2}
        result = _apply_decimal_config(3.14159, "4", config)
        assert result == "3.14"
    
    def test_missing_column_key(self):
        """测试缺失列键时使用默认值"""
        config = {"5": 2}
        result = _apply_decimal_config(3.14159, "4", config)
        assert result == "3.1"  # 默认1位小数
    
    def test_empty_config(self):
        """测试空配置"""
        result = _apply_decimal_config(3.14159, "4", {})
        assert result == "3.1"
    
    def test_conditional_config_greater_equal(self):
        """测试条件配置 >="""
        config = {
            "4": {
                "condition": ">= 100",
                "true": 0,
                "false": 1
            }
        }
        # 值 >= 100
        result = _apply_decimal_config(150.5, "4", config)
        assert result == "150"  # 0位小数（舍去小数部分）
        # 值 < 100
        result = _apply_decimal_config(50.5, "4", config)
        assert result == "50.5"  # 1位小数
    
    def test_conditional_config_greater(self):
        """测试条件配置 >"""
        config = {
            "4": {
                "condition": "> 100",
                "true": 0,
                "false": 2
            }
        }
        result = _apply_decimal_config(100.1, "4", config)
        assert result == "100"  # 大于100
        result = _apply_decimal_config(100.0, "4", config)
        assert result == "100.00"  # 不大于100
    
    def test_conditional_config_less_equal(self):
        """测试条件配置 <="""
        config = {
            "4": {
                "condition": "<= 50",
                "true": 3,
                "false": 1
            }
        }
        result = _apply_decimal_config(50.0, "4", config)
        assert result == "50.000"
        result = _apply_decimal_config(51.0, "4", config)
        assert result == "51.0"
    
    def test_conditional_config_less(self):
        """测试条件配置 <"""
        config = {
            "4": {
                "condition": "< 100",
                "true": 2,
                "false": 0
            }
        }
        result = _apply_decimal_config(99.9, "4", config)
        assert result == "99.90"
        result = _apply_decimal_config(100.0, "4", config)
        assert result == "100"
    
    def test_conditional_config_equal(self):
        """测试条件配置 =="""
        config = {
            "4": {
                "condition": "== 100",
                "true": 0,
                "false": 2
            }
        }
        result = _apply_decimal_config(100.0, "4", config)
        assert result == "100"
        result = _apply_decimal_config(99.9, "4", config)
        assert result == "99.90"
    
    def test_invalid_condition_format(self):
        """测试无效的条件格式"""
        config = {
            "4": {
                "condition": "invalid",
                "true": 0,
                "false": 1
            }
        }
        result = _apply_decimal_config(100.0, "4", config)
        assert result == "100.0"  # 使用false_decimal


# =============================================================================
# calculate_rated_energy_efficacy 函数测试
# =============================================================================

class TestCalculateRatedEnergyEfficacy:
    """calculate_rated_energy_efficacy 函数的单元测试"""
    
    def test_basic_calculation(self):
        """测试基本计算"""
        photometric_data = [
            [0, 10.0, 0, 1000.0],  # 功率10，光通1000
            [0, 20.0, 0, 2000.0],  # 功率20，光通2000
        ]
        # 平均功率 = 15, 平均光通 = 1500, 能效 = 1500/15 = 100
        result = calculate_rated_energy_efficacy(photometric_data)
        assert result == "100.0"
    
    def test_custom_digits(self):
        """测试自定义小数位"""
        photometric_data = [
            [0, 10.0, 0, 1000.0],
        ]
        result = calculate_rated_energy_efficacy(photometric_data, digits=2)
        assert result == "100.00"
    
    def test_empty_table(self):
        """测试空表"""
        assert calculate_rated_energy_efficacy([]) == "N/A"
        assert calculate_rated_energy_efficacy(None) == "N/A"
    
    def test_with_dict_input(self):
        """测试字典格式输入"""
        photometric_data = {
            'value': [[0, 10.0, 0, 1000.0]]
        }
        result = calculate_rated_energy_efficacy(photometric_data)
        assert result == "100.0"
    
    def test_missing_columns(self):
        """测试缺少列的情况"""
        photometric_data = [
            [0],  # 缺少第2列和第4列
            [0, 10.0],  # 缺少第4列
        ]
        result = calculate_rated_energy_efficacy(photometric_data)
        assert result == "N/A"
    
    def test_invalid_values(self):
        """测试无效值"""
        photometric_data = [
            [0, "invalid", 0, "value"],
        ]
        result = calculate_rated_energy_efficacy(photometric_data)
        assert result == "N/A"
    
    def test_zero_average_power(self):
        """测试平均功率为零"""
        photometric_data = [
            [0, 0.0, 0, 1000.0],
        ]
        result = calculate_rated_energy_efficacy(photometric_data)
        assert result == "N/A"


# =============================================================================
# calculate_rated_energy_class_rating 函数测试
# =============================================================================

class TestCalculateRatedEnergyClassRating:
    """calculate_rated_energy_class_rating 函数的单元测试"""
    
    def test_class_a_plus_plus(self):
        """测试 A++ 等级 (η_TM >= 210)"""
        # 需要 η_TM >= 210, 即 flux/power >= 210 (NDLS+MLS)
        photometric_data = [
            [0, 10.0, 0, 2200.0],  # 220 lm/W
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "2200", True, True
        )
        assert result == "A"
    
    def test_class_a(self):
        """测试 A 等级"""
        # η_TM >= 210 返回 A
        # 对于 NDLS+MLS (F_TM=1.0), 需要 flux/power >= 210
        photometric_data = [
            [0, 10.0, 0, 2100.0],  # 210 lm/W
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "2100", True, True
        )
        assert result == "A"
    
    def test_class_b(self):
        """测试 B 等级"""
        # 185 <= η_TM < 210 返回 B
        photometric_data = [
            [0, 10.0, 0, 1900.0],  # 190 lm/W
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "1900", True, True
        )
        assert result == "B"
    
    def test_class_c(self):
        """测试 C 等级"""
        # 160 <= η_TM < 185 返回 C
        photometric_data = [
            [0, 10.0, 0, 1700.0],  # 170 lm/W
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "1700", True, True
        )
        assert result == "C"
    
    def test_class_d(self):
        """测试 D 等级"""
        # 135 <= η_TM < 160 返回 D
        photometric_data = [
            [0, 10.0, 0, 1400.0],  # 140 lm/W
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "1400", True, True
        )
        assert result == "D"
    
    def test_class_e(self):
        """测试 E 等级"""
        # 110 <= η_TM < 135 返回 E
        photometric_data = [
            [0, 10.0, 0, 1200.0],  # 120 lm/W
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "1200", True, True
        )
        assert result == "E"
    
    def test_class_f(self):
        """测试 F 等级"""
        # 85 <= η_TM < 110 返回 F
        photometric_data = [
            [0, 10.0, 0, 900.0],  # 90 lm/W
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "900", True, True
        )
        assert result == "F"
    
    def test_class_g(self):
        """测试 G 等级"""
        photometric_data = [
            [0, 10.0, 0, 500.0],  # 50 lm/W
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "500", True, True
        )
        assert result == "G"
    
    def test_dls_mls(self):
        """测试 DLS + MLS (F_TM = 1.176)"""
        photometric_data = [
            [0, 10.0, 0, 1000.0],  # 100 lm/W base
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "1000", False, True
        )
        # η_TM = 100 * 1.176 = 117.6, should be E class (110 <= 117.6 < 135)
        assert result == "E"
    
    def test_ndls_nmls(self):
        """测试 NDLS + NMLS (F_TM = 0.926)"""
        photometric_data = [
            [0, 10.0, 0, 1500.0],  # 150 lm/W base
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "1500", True, False
        )
        # η_TM = 150 * 0.926 = 138.9, should be D class
        assert result == "D"
    
    def test_dls_nmls(self):
        """测试 DLS + NMLS (F_TM = 1.089)"""
        photometric_data = [
            [0, 10.0, 0, 1000.0],
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "10", "1000", False, False
        )
        # η_TM = 100 * 1.089 = 108.9, should be F class (85 <= 108.9 < 110)
        assert result == "F"
    
    def test_empty_table(self):
        """测试空表"""
        result = calculate_rated_energy_class_rating([], "10", "1000", True, True)
        assert result == "N/A"
    
    def test_zero_power(self):
        """测试功率为零"""
        photometric_data = [
            [0, 0.0, 0, 1000.0],
        ]
        result = calculate_rated_energy_class_rating(
            photometric_data, "0", "1000", True, True
        )
        assert result == "N/A"


# =============================================================================
# format_sample_size 函数测试
# =============================================================================

class TestFormatSampleSize:
    """format_sample_size 函数的单元测试"""
    
    def test_with_controlgear(self):
        """测试有控制装置型号"""
        result = format_sample_size("product", "light", "CG-001")
        assert result == "10 pcs per model + 3 pcs controlgear"
    
    def test_without_controlgear(self):
        """测试无控制装置型号"""
        result = format_sample_size("product", "light", "")
        assert result == "10 pcs per model"
    
    def test_with_none_controlgear(self):
        """测试控制装置型号为 None"""
        result = format_sample_size("product", "light", None)
        assert result == "10 pcs per model"


# =============================================================================
# calculate_directional_info 函数测试
# =============================================================================

class TestCalculateDirectionalInfo:
    """calculate_directional_info 函数的单元测试"""
    
    def test_non_directional_true(self):
        """测试非定向为 True"""
        result = calculate_directional_info({"type": "checkbox", "value": "true"})
        assert result == "non-directional"
    
    def test_non_directional_false(self):
        """测试非定向为 False"""
        result = calculate_directional_info({"type": "checkbox", "value": "false"})
        assert result == "directional"
    
    def test_none_input(self):
        """测试 None 输入"""
        result = calculate_directional_info(None)
        assert result == "unknown"
    
    def test_invalid_dict(self):
        """测试无效字典"""
        result = calculate_directional_info({"other": "value"})
        assert result == "unknown"


# =============================================================================
# calculate_ponmax 函数测试
# =============================================================================

class TestCalculatePonmax:
    """calculate_ponmax 函数的单元测试"""
    
    def test_ndls_basic(self):
        """测试 NDLS 基本计算"""
        # Ponmax = C × (L + Φuse/(F × η)) × R
        # C=1.0, L=2.0, F=1.00, η=120, R=(80+80)/160=1.0
        # Ponmax = 1.0 * (2.0 + 1000/(1.0*120)) * 1.0 = 2 + 8.33 = 10.33
        result = calculate_ponmax(
            useful_luminous_flux="1000",
            non_directional=True,
            cri="80",
            led_source=True,
            color_tuneable=False,
            mains=True
        )
        assert result == "10.33"
    
    def test_dls_basic(self):
        """测试 DLS 基本计算"""
        # C=0.79, L=2.0, F=0.85, η=120, R=1.0
        # Ponmax = 0.79 * (2.0 + 1000/(0.85*120)) * 1.0
        result = calculate_ponmax(
            useful_luminous_flux="1000",
            non_directional=False,
            cri="80",
            led_source=True,
            color_tuneable=False,
            mains=True
        )
        expected = 0.79 * (2.0 + 1000/(0.85*120)) * 1.0
        assert float(result) == pytest.approx(expected, rel=0.01)
    
    def test_color_tuneable(self):
        """测试可调色光源"""
        # 可调色光源强制 C=1.0
        result = calculate_ponmax(
            useful_luminous_flux="1000",
            non_directional=False,  # DLS
            cri="80",
            led_source=True,
            color_tuneable=True,  # 可调色
            mains=True
        )
        # 可调色时C=1.0
        expected = 1.0 * (2.0 + 1000/(0.85*120)) * 1.0
        assert float(result) == pytest.approx(expected, rel=0.01)
    
    def test_low_cri(self):
        """测试低 CRI (<= 25)"""
        # R = 0.65 for CRI <= 25
        result = calculate_ponmax(
            useful_luminous_flux="1000",
            non_directional=True,
            cri="20",
            led_source=True,
            color_tuneable=False,
            mains=True
        )
        expected = 1.0 * (2.0 + 1000/120) * 0.65
        assert float(result) == pytest.approx(expected, rel=0.01)
    
    def test_high_cri(self):
        """测试高 CRI"""
        # R = (CRI + 80) / 160
        # CRI=90: R = 170/160 = 1.0625
        result = calculate_ponmax(
            useful_luminous_flux="1000",
            non_directional=True,
            cri="90",
            led_source=True,
            color_tuneable=False,
            mains=True
        )
        expected = 1.0 * (2.0 + 1000/120) * 1.0625
        assert float(result) == pytest.approx(expected, rel=0.01)
    
    def test_checkbox_input(self):
        """测试 checkbox 格式输入"""
        result = calculate_ponmax(
            useful_luminous_flux="1000",
            non_directional={"type": "checkbox", "value": "true"},
            cri="80",
            led_source={"type": "checkbox", "value": "true"},
            color_tuneable={"type": "checkbox", "value": "false"},
            mains={"type": "checkbox", "value": "true"}
        )
        expected = 1.0 * (2.0 + 1000/120) * 1.0
        assert float(result) == pytest.approx(expected, rel=0.01)
    
    def test_invalid_flux(self):
        """测试无效光通量"""
        result = calculate_ponmax(
            useful_luminous_flux="invalid",
            non_directional=True,
            cri="80",
            led_source=True,
            color_tuneable=False,
            mains=True
        )
        assert result == "N/A"
    
    def test_none_flux(self):
        """测试 None 光通量"""
        # None 被转为 0, 所以计算结果 = C * (L + 0) * R = 1.0 * 2.0 * 1.0 = 2.0
        result = calculate_ponmax(
            useful_luminous_flux=None,
            non_directional=True,
            cri="80",
            led_source=True,
            color_tuneable=False,
            mains=True
        )
        assert result == "2.00"


# =============================================================================
# _parse_checkbox_value 函数测试
# =============================================================================

class TestParseCheckboxValue:
    """_parse_checkbox_value 函数的单元测试"""
    
    def test_true_boolean(self):
        """测试 True 布尔值"""
        assert _parse_checkbox_value(True) is True
    
    def test_false_boolean(self):
        """测试 False 布尔值"""
        assert _parse_checkbox_value(False) is False
    
    def test_true_string(self):
        """测试 "true" 字符串"""
        assert _parse_checkbox_value("true") is True
        assert _parse_checkbox_value("True") is True
        assert _parse_checkbox_value("TRUE") is True
    
    def test_false_string(self):
        """测试 "false" 字符串"""
        assert _parse_checkbox_value("false") is False
        assert _parse_checkbox_value("False") is False
    
    def test_yes_string(self):
        """测试 "yes" 字符串"""
        assert _parse_checkbox_value("yes") is True
        assert _parse_checkbox_value("YES") is True
    
    def test_one_string(self):
        """测试 "1" 字符串"""
        assert _parse_checkbox_value("1") is True
    
    def test_checkbox_dict_true(self):
        """测试 checkbox 字典 True"""
        assert _parse_checkbox_value({"type": "checkbox", "value": "true"}) is True
    
    def test_checkbox_dict_false(self):
        """测试 checkbox 字典 False"""
        assert _parse_checkbox_value({"type": "checkbox", "value": "false"}) is False
    
    def test_value_dict_true(self):
        """测试带 value 键的字典 True"""
        assert _parse_checkbox_value({"value": "true"}) is True
    
    def test_value_dict_false(self):
        """测试带 value 键的字典 False"""
        assert _parse_checkbox_value({"value": "false"}) is False
    
    def test_value_dict_numeric(self):
        """测试带 value 键的字典数值"""
        assert _parse_checkbox_value({"value": 1}) is True
        assert _parse_checkbox_value({"value": 0}) is False
    
    def test_none_input(self):
        """测试 None 输入"""
        assert _parse_checkbox_value(None) is False
    
    def test_empty_string(self):
        """测试空字符串"""
        assert _parse_checkbox_value("") is False
    
    def test_number_true(self):
        """测试非零数字"""
        assert _parse_checkbox_value(1) is True
        assert _parse_checkbox_value(42) is True
    
    def test_number_false(self):
        """测试零数字"""
        assert _parse_checkbox_value(0) is False


# =============================================================================
# calculate_required_maintenance_percentage 函数测试
# =============================================================================

class TestCalculateRequiredMaintenancePercentage:
    """calculate_required_maintenance_percentage 函数的单元测试"""
    
    def test_basic_calculation(self):
        """测试基本计算"""
        # XLMF,MIN % = 100 × e^((3000 × ln(0.7)) / L70)
        # L70 = 25000: XLMF = 100 * exp(3000 * ln(0.7) / 25000)
        l70 = 25000
        expected = 100 * math.exp((3000 * math.log(0.7)) / l70)
        result = calculate_required_maintenance_percentage(l70)
        assert result == f"{expected:.1f}%"
    
    def test_with_dict_input(self):
        """测试字典格式输入"""
        result = calculate_required_maintenance_percentage({"value": "25000"})
        expected = 100 * math.exp((3000 * math.log(0.7)) / 25000)
        assert result == f"{expected:.1f}%"
    
    def test_low_l70_uses_default(self):
        """测试低 L70 使用默认值"""
        # L70 < 1000 时使用默认值 25000
        result = calculate_required_maintenance_percentage("500")
        expected = 100 * math.exp((3000 * math.log(0.7)) / 25000)
        assert result == f"{expected:.1f}%"
    
    def test_cap_at_96_percent(self):
        """测试上限为 96%"""
        # 极高的 L70 值应该返回 96.0%
        result = calculate_required_maintenance_percentage("100000")
        assert result == "96.0%"
    
    def test_invalid_l70(self):
        """测试无效 L70"""
        assert calculate_required_maintenance_percentage("invalid") == "N/A"
    
    def test_zero_l70(self):
        """测试零 L70 - 使用默认值 25000"""
        # 0 < 1000, 所以使用默认值 25000
        result = calculate_required_maintenance_percentage("0")
        expected = 100 * math.exp((3000 * math.log(0.7)) / 25000)
        assert result == f"{expected:.1f}%"
    
    def test_negative_l70(self):
        """测试负 L70 - 使用默认值 25000"""
        # -1000 < 1000, 所以使用默认值 25000
        result = calculate_required_maintenance_percentage("-1000")
        expected = 100 * math.exp((3000 * math.log(0.7)) / 25000)
        assert result == f"{expected:.1f}%"


# =============================================================================
# calculate_light_source_tech 函数测试
# =============================================================================

class TestCalculateLightSourceTech:
    """calculate_light_source_tech 函数的单元测试"""
    
    def test_led_true(self):
        """测试 LED 为 True"""
        result = calculate_light_source_tech({"type": "checkbox", "value": "true"})
        assert result == "LED"
    
    def test_led_false(self):
        """测试 LED 为 False"""
        result = calculate_light_source_tech({"type": "checkbox", "value": "false"})
        assert result == "non-LED"
    
    def test_none_input(self):
        """测试 None 输入"""
        result = calculate_light_source_tech(None)
        assert result == "unknown"
    
    def test_invalid_dict(self):
        """测试无效字典"""
        result = calculate_light_source_tech({"other": "value"})
        assert result == "unknown"


# =============================================================================
# calculated_zone_table 函数测试
# =============================================================================

class TestCalculatedZoneTable:
    """calculated_zone_table 函数的单元测试"""
    
    def test_beam_angle_30(self):
        """测试光束角 <= 30"""
        result = calculated_zone_table(100, 200, 300, 400, 500, "25")
        assert result == [["zone-0_30", 100]]
    
    def test_beam_angle_60(self):
        """测试光束角 <= 60"""
        result = calculated_zone_table(100, 200, 300, 400, 500, "45")
        assert result == [["zone-0_30", 100], ["zone-0_60", 200]]
    
    def test_beam_angle_90(self):
        """测试光束角 <= 90"""
        result = calculated_zone_table(100, 200, 300, 400, 500, "85")
        assert result == [
            ["zone-0_30", 100],
            ["zone-0_60", 200],
            ["zone-0_90", 300]
        ]
    
    def test_beam_angle_120(self):
        """测试光束角 <= 120"""
        result = calculated_zone_table(100, 200, 300, 400, 500, "100")
        assert result == [
            ["zone-0_30", 100],
            ["zone-0_60", 200],
            ["zone-0_90", 300],
            ["zone-0_120", 400]
        ]
    
    def test_beam_angle_180(self):
        """测试光束角 > 120"""
        result = calculated_zone_table(100, 200, 300, 400, 500, "150")
        assert result == [
            ["zone-0_30", 100],
            ["zone-0_60", 200],
            ["zone-0_90", 300],
            ["zone-0_120", 400],
            ["zone-180", 500]
        ]
    
    def test_none_beam_angle(self):
        """测试 None 光束角"""
        result = calculated_zone_table(100, 200, 300, 400, 500, None)
        assert result is None
    
    def test_invalid_beam_angle(self):
        """测试无效光束角"""
        result = calculated_zone_table(100, 200, 300, 400, 500, "invalid")
        assert result is None


# =============================================================================
# calculated_beam_table 函数测试
# =============================================================================

class TestCalculatedBeamTable:
    """calculated_beam_table 函数的单元测试"""
    
    def test_basic(self):
        """测试基本返回"""
        result = calculated_beam_table("90", "1000")
        assert result == [["beam angle", "90"], ["peak intensity", "1000"]]
    
    def test_with_numbers(self):
        """测试数值输入"""
        result = calculated_beam_table(90, 1000)
        assert result == [["beam angle", 90], ["peak intensity", 1000]]


# =============================================================================
# calculate_tdb_remarks 函数测试
# =============================================================================

class TestCalculateTdbRemarks:
    """calculate_tdb_remarks 函数的单元测试"""
    
    def test_initial_test(self):
        """测试初始测试"""
        result = calculate_tdb_remarks({"type": "checkbox", "value": "true"})
        assert result == "except 3600hrs Lumen Maintenance test"
    
    def test_not_initial_test(self):
        """测试非初始测试"""
        result = calculate_tdb_remarks({"type": "checkbox", "value": "false"})
        assert result == ""
    
    def test_boolean_true(self):
        """测试布尔值 True"""
        result = calculate_tdb_remarks(True)
        assert result == "except 3600hrs Lumen Maintenance test"
    
    def test_boolean_false(self):
        """测试布尔值 False"""
        result = calculate_tdb_remarks(False)
        assert result == ""
    
    def test_string_true(self):
        """测试字符串 true"""
        result = calculate_tdb_remarks("true")
        assert result == "except 3600hrs Lumen Maintenance test"
    
    def test_string_false(self):
        """测试字符串 false"""
        result = calculate_tdb_remarks("false")
        assert result == ""
