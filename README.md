重复阈值（threshold）是指 图片需要至少在多少个不同的Excel文件中出现 才会被判定为重复图片。例如：

阈值	含义	典型应用场景
1	所有图片都会被标记为重复（无效设置）	测试时使用
2	图片在2个及以上文件中出现时标记为重复	精确查找重复
3	图片在3个及以上文件中出现时标记为重复	排除偶然重复
5	图片在5个及以上文件中出现时标记为重复	查找高频重复
≥6	更高要求的重复检测	特殊场景需求
阈值设置建议
1. 推荐默认值：5
适用场景：常规查重需求

效果：过滤偶发重复，聚焦高频重复图片

2. 低阈值（2-3）
适用场景：

需要查找所有可能的重复

文件数量较少（<10个Excel文件）

风险：可能包含偶然重复（如通用LOGO）

3. 高阈值（>5）
适用场景：

大规模文件集（>100个Excel）

查找系统性重复（如模板图片）

风险：可能漏检低频重复

示例场景分析
假设有 10个Excel文件，其中某图片出现情况：

图片哈希	出现文件数	不同阈值的检测结果
A1B2C3	8	阈值≥2时均会检出
D4E5F6	5	阈值≤5时检出
G7H8I9	3	阈值≤3时检出
J0K1L2	1	不检出
技术实现逻辑
在代码中通过以下逻辑实现阈值过滤：


最佳实践建议
首次使用：从默认阈值5开始，根据结果调整

精确查找：设置为2，但需人工复核结果

批量处理：根据文件总量动态调整（阈值 = 文件总数×10%）

通过合理设置阈值，可以平衡 查全率 和 查准率，高效定位目标重复图片
