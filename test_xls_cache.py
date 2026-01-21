"""
测试 .xls 文件转换缓存性能
这个脚本展示缓存机制如何提高重复操作的速度
"""
import time
import os
from src.nodes.merge_nodes import convert_xls_to_xlsx, clear_xls_conversion_cache

def test_conversion_speed(xls_file_path, iterations=3):
    """测试转换速度"""
    if not os.path.exists(xls_file_path):
        print(f"错误: 找不到文件 {xls_file_path}")
        return
    
    print("=" * 60)
    print(f"测试文件: {xls_file_path}")
    print(f"测试迭代次数: {iterations}")
    print("=" * 60)
    
    times = []
    
    for i in range(iterations):
        print(f"\n第 {i+1} 次转换:")
        start = time.time()
        try:
            temp_path = convert_xls_to_xlsx(xls_file_path, use_cache=True)
            elapsed = time.time() - start
            times.append(elapsed)
            print(f"总耗时: {elapsed:.2f}秒")
            print(f"临时文件: {temp_path}")
        except Exception as e:
            print(f"转换失败: {e}")
            return
    
    print("\n" + "=" * 60)
    print("性能统计:")
    print(f"  第1次 (无缓存): {times[0]:.2f}秒")
    if len(times) > 1:
        avg_cached = sum(times[1:]) / len(times[1:])
        print(f"  后续平均 (有缓存): {avg_cached:.2f}秒")
        speedup = times[0] / avg_cached
        print(f"  速度提升: {speedup:.1f}x")
    print("=" * 60)
    
    # 清理缓存
    print("\n清理缓存...")
    clear_xls_conversion_cache()

if __name__ == "__main__":
    # 使用示例
    # 请将下面的路径替换为你的 .xls 文件路径
    test_file = r"C:\path\to\your\file.xls"
    
    if os.path.exists(test_file):
        test_conversion_speed(test_file, iterations=3)
    else:
        print("请在脚本中设置正确的 .xls 文件路径")
        print("\n使用方法:")
        print("1. 找到一个 .xls 文件")
        print("2. 修改 test_file 变量为该文件的完整路径")
        print("3. 运行此脚本")
        print("\n预期效果:")
        print("- 第1次转换会花费较长时间（实际转换）")
        print("- 后续转换会非常快（使用缓存，几乎瞬间完成）")
