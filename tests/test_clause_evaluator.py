"""
General Clause Evaluator Test Suite
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.clause_evaluator import evaluate_clause, ConditionEvaluator


def test_basic_comparison():
    """Test basic comparison"""
    print("\n=== Test Basic Comparison ===")
    
    clause_config = {
        "clause_id": "test_basic",
        "param_names": ["value"],
        "rules": [
            {"condition": "value > 100", "result": "High"},
            {"condition": "value > 50", "result": "Medium"}
        ],
        "default": "Low"
    }
    
    test_cases = [
        (150, "High"),
        (75, "Medium"),
        (25, "Low")
    ]
    
    for value, expected in test_cases:
        result = evaluate_clause(value, clause_config=clause_config)
        status = "[OK]" if result == expected else "[FAIL]"
        print(f"{status} value={value} => {result} (expected: {expected})")


def test_logical_operators():
    """Test logical operators"""
    print("\n=== Test Logical Operators ===")
    
    clause_config = {
        "clause_id": "test_logical",
        "param_names": ["A", "B", "C"],
        "rules": [
            {"condition": "A == 'true' AND B == 'true'", "result": "Both True"},
            {"condition": "A == 'true' OR C == 'true'", "result": "At Least One True"},
            {"condition": "NOT (A == 'true')", "result": "A is False"}
        ],
        "default": "Default"
    }
    
    test_cases = [
        (("true", "true", "false"), "Both True"),
        (("true", "false", "true"), "At Least One True"),
        (("false", "false", "true"), "At Least One True"),
        (("false", "true", "false"), "A is False"),
    ]
    
    for args, expected in test_cases:
        result = evaluate_clause(*args, clause_config=clause_config)
        status = "[OK]" if result == expected else "[FAIL]"
        print(f"{status} {args} => {result} (expected: {expected})")


def test_in_operator():
    """Test IN operator"""
    print("\n=== Test IN Operator ===")
    
    clause_config = {
        "clause_id": "test_in",
        "param_names": ["category"],
        "rules": [
            {"condition": "category IN ['A', 'B', 'C']", "result": "Standard"},
            {"condition": "category IN ['X', 'Y']", "result": "Special"},
            {"condition": "category NOT IN ['A', 'B', 'C', 'X', 'Y']", "result": "Other"}
        ],
        "default": "Unknown"
    }
    
    test_cases = [
        ("A", "Standard"),
        ("X", "Special"),
        ("Z", "Other"),
        ("D", "Other")
    ]
    
    for category, expected in test_cases:
        result = evaluate_clause(category, clause_config=clause_config)
        status = "[OK]" if result == expected else "[FAIL]"
        print(f"{status} category={category} => {result} (expected: {expected})")


def test_complex_case():
    """Test complex clause evaluation"""
    print("\n=== Test Complex Case - Clause Evaluation ===")
    
    clause_config = {
        "clause_id": "v.1.a",
        "param_names": ["containing_product", "light_sources"],
        "rules": [
            {
                "condition": "containing_product == 'true' AND light_sources == 'true'",
                "result": "Pass"
            },
            {
                "condition": "containing_product == 'false'",
                "result": "N/A"
            }
        ],
        "default": "Fail"
    }
    
    test_cases = [
        (("true", "true"), "Pass"),
        (("true", "false"), "Fail"),
        (("false", "true"), "N/A"),
        (("false", "false"), "N/A"),
    ]
    
    for args, expected in test_cases:
        result = evaluate_clause(*args, clause_config=clause_config)
        status = "[OK]" if result == expected else "[FAIL]"
        cp, ls = args
        print(f"{status} containing_product={cp}, light_sources={ls} => {result} (expected: {expected})")


def run_all_tests():
    """Run all tests"""
    print("=" * 70)
    print("General Clause Evaluator Test Suite")
    print("=" * 70)
    
    test_basic_comparison()
    test_logical_operators()
    test_in_operator()
    test_complex_case()
    
    print("\n" + "=" * 70)
    print("All tests completed!")
    print("=" * 70)


if __name__ == "__main__":
    run_all_tests()
