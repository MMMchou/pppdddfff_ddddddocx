#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PaddleOCR PDF 转 Word/Markdown 工具
使用 paddleocr 命令行工具
"""

import os
import sys
import argparse
import subprocess
from pathlib import Path
from typing import List
import logging
from tqdm import tqdm

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)

CURRENT_DIR = Path(__file__).parent.absolute()


def convert_pdf(pdf_path: str, output_dir: str = None, use_gpu: bool = False,
               enable_table: bool = True) -> dict:
    """
    转换单个 PDF 文件
    
    Args:
        pdf_path: PDF 文件路径
        output_dir: 输出目录
        use_gpu: 是否使用 GPU
        enable_table: 是否启用表格识别
        
    Returns:
        转换结果字典
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        return {'status': 'failed', 'input': str(pdf_path), 'error': '文件不存在'}
    
    # 准备输出目录
    if output_dir is None:
        output_dir = CURRENT_DIR / 'output' / pdf_path.stem
    else:
        output_dir = Path(output_dir) / pdf_path.stem
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"转换: {pdf_path.name} -> {output_dir}")
    
    # 构建命令 (使用 PaddleOCR 3.x 的命令格式)
    cmd = [
        'paddleocr',
        'pp_structurev3',  # 使用 PP-Structure V3
        '--input', str(pdf_path.absolute()),
        '--save_path', str(output_dir.absolute()),
        '--device', 'gpu' if use_gpu else 'cpu',
    ]
    
    # 表格识别选项
    if enable_table:
        cmd.extend(['--use_table_recognition', 'True'])
    
    try:
        # 执行命令
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        
        if result.returncode == 0:
            # 整理输出文件
            logger.info("正在整理输出文件...")
            from organize_output import organize_output_directory
            organize_output_directory(str(output_dir))
            
            # 查找最终文件
            final_dir = output_dir / 'final'
            outputs = {}
            
            final_docx = final_dir / f"{pdf_path.stem}.docx"
            if final_docx.exists():
                outputs['docx'] = str(final_docx)
                logger.info(f"✓ Word 文档: {final_docx}")
            
            final_md = final_dir / f"{pdf_path.stem}.md"
            if final_md.exists():
                outputs['markdown'] = str(final_md)
                logger.info(f"✓ Markdown 文档: {final_md}")
            
            return {
                'status': 'success',
                'input': str(pdf_path),
                'outputs': outputs,
                'output_dir': str(output_dir)
            }
        else:
            return {
                'status': 'failed',
                'input': str(pdf_path),
                'error': result.stderr or '转换失败'
            }
    
    except Exception as e:
        return {
            'status': 'failed',
            'input': str(pdf_path),
            'error': str(e)
        }


def convert_batch(input_dir: str, output_dir: str = None, use_gpu: bool = False,
                 enable_table: bool = True) -> List[dict]:
    """批量转换 PDF 文件"""
    input_path = Path(input_dir)
    if not input_path.exists():
        logger.error(f"目录不存在: {input_dir}")
        return []
    
    pdf_files = list(input_path.rglob("*.pdf"))
    
    if len(pdf_files) == 0:
        logger.warning(f"未找到 PDF 文件: {input_dir}")
        return []
    
    logger.info(f"找到 {len(pdf_files)} 个 PDF 文件")
    
    results = []
    for pdf_file in tqdm(pdf_files, desc="转换进度", ncols=80):
        result = convert_pdf(str(pdf_file), output_dir, use_gpu, enable_table)
        results.append(result)
    
    success = sum(1 for r in results if r['status'] == 'success')
    failed = len(results) - success
    
    logger.info(f"完成！成功: {success}, 失败: {failed}")
    
    return results


def main():
    parser = argparse.ArgumentParser(
        description='PaddleOCR PDF 转 Word/Markdown 工具',
        epilog='示例: python convert.py input.pdf'
    )
    
    parser.add_argument('input', help='PDF 文件或目录')
    parser.add_argument('-o', '--output', help='输出目录')
    parser.add_argument('--batch', action='store_true', help='批量处理')
    parser.add_argument('--no-table', action='store_true', help='禁用表格识别')
    parser.add_argument('--gpu', action='store_true', help='使用 GPU')
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    
    if not input_path.exists():
        logger.error(f"路径不存在: {args.input}")
        sys.exit(1)
    
    if args.batch or input_path.is_dir():
        # 批量转换
        results = convert_batch(
            str(input_path),
            args.output,
            args.gpu,
            not args.no_table
        )
        
        # 保存摘要
        output_base = Path(args.output) if args.output else CURRENT_DIR / 'output'
        summary_file = output_base / 'summary.txt'
        summary_file.parent.mkdir(parents=True, exist_ok=True)
        
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write("转换结果摘要\n" + "=" * 60 + "\n\n")
            for r in results:
                f.write(f"文件: {r['input']}\n")
                f.write(f"状态: {r['status']}\n")
                if r['status'] == 'success':
                    for k, v in r.get('outputs', {}).items():
                        f.write(f"  {k}: {v}\n")
                else:
                    f.write(f"  错误: {r.get('error')}\n")
                f.write("\n")
        
        print(f"\n摘要已保存: {summary_file}")
    
    else:
        # 单文件转换
        result = convert_pdf(
            str(input_path),
            args.output,
            args.gpu,
            not args.no_table
        )
        
        if result['status'] == 'success':
            print(f"\n✓ 转换成功！")
            for k, v in result.get('outputs', {}).items():
                print(f"  {k}: {v}")
        else:
            print(f"\n✗ 失败: {result.get('error')}")
            sys.exit(1)


if __name__ == '__main__':
    main()

