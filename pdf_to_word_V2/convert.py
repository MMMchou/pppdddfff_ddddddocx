#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 转 DOCX 批量转换脚本
基于 pdf2docx 库实现
"""

import argparse
import logging
import time
from pathlib import Path
from typing import Optional, Dict, Any
import yaml

try:
    from pdf2docx import Converter  # pyright: ignore[reportMissingImports]
except ImportError:
    print("错误: 未安装 pdf2docx 库")
    print("请运行: pip install -r requirements.txt")
    exit(1)


class PDFConverter:
    """PDF 到 DOCX 转换器"""
    
    def __init__(self, config_path: str = "config.yaml"):
        """
        初始化转换器
        
        Args:
            config_path: 配置文件路径
        """
        self.config = self._load_config(config_path)
        self._setup_logging()
        
    def _load_config(self, config_path: str) -> Dict[str, Any]:
        """加载配置文件"""
        config_file = Path(config_path)
        if not config_file.exists():
            logging.warning(f"配置文件 {config_path} 不存在，使用默认配置")
            return self._default_config()
        
        with open(config_file, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    
    def _default_config(self) -> Dict[str, Any]:
        """返回默认配置"""
        return {
            'input_dir': 'pdf_data',
            'output_dir': 'output',
            'conversion': {
                'parse_lattice_table': True,
                'multi_processing': False
            },
            'debug': {
                'enable': False,
                'verbose': True
            },
            'error_handling': {
                'enable_fallback': True,
                'continue_on_error': True
            }
        }
    
    def _setup_logging(self):
        """设置日志"""
        level = logging.INFO if self.config['debug']['verbose'] else logging.WARNING
        logging.basicConfig(
            level=level,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        self.logger = logging.getLogger(__name__)
    
    def convert_single(
        self,
        pdf_path: Path,
        output_dir: Path,
        enable_debug: bool = False,
        settings_override: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        转换单个 PDF 文件
        
        Args:
            pdf_path: PDF 文件路径
            output_dir: 输出目录
            enable_debug: 是否启用调试模式
            settings_override: 覆盖配置参数
            
        Returns:
            转换结果字典，包含 success, message, output_path, use_fallback 等信息
        """
        result = {
            'success': False,
            'message': '',
            'output_path': None,
            'use_fallback': False,
            'duration': 0
        }
        
        start_time = time.time()
        
        # 创建输出目录
        file_output_dir = output_dir / pdf_path.stem
        file_output_dir.mkdir(parents=True, exist_ok=True)
        
        # 设置输出路径
        docx_path = file_output_dir / f"{pdf_path.stem}.docx"
        log_path = file_output_dir / "conversion.log"
        
        # 设置文件日志
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.logger.addHandler(file_handler)
        
        try:
            self.logger.info(f"开始转换: {pdf_path.name}")
            
            # 准备转换参数
            kwargs = settings_override or self.config['conversion'].copy()
            
            # 首次尝试转换
            success = self._do_convert(pdf_path, docx_path, kwargs, enable_debug, file_output_dir)
            
            if not success and self.config['error_handling']['enable_fallback']:
                # 使用 fallback 配置重试
                self.logger.warning(f"标准配置转换失败，尝试 fallback 模式（关闭 lattice 表格解析）")
                kwargs['parse_lattice_table'] = False
                success = self._do_convert(pdf_path, docx_path, kwargs, enable_debug, file_output_dir)
                if success:
                    result['use_fallback'] = True
            
            if success:
                result['success'] = True
                result['output_path'] = str(docx_path)
                result['message'] = '转换成功'
                if result['use_fallback']:
                    result['message'] += ' (使用 fallback 配置)'
                self.logger.info(f"✓ 转换成功: {pdf_path.name} -> {docx_path.name}")
            else:
                result['message'] = '转换失败'
                self.logger.error(f"✗ 转换失败: {pdf_path.name}")
                
        except Exception as e:
            result['message'] = f'转换异常: {str(e)}'
            self.logger.exception(f"转换异常: {pdf_path.name}: {e}")
        finally:
            result['duration'] = time.time() - start_time
            self.logger.removeHandler(file_handler)
            file_handler.close()
        
        return result
    
    def _do_convert(
        self,
        pdf_path: Path,
        docx_path: Path,
        kwargs: Dict[str, Any],
        enable_debug: bool,
        output_dir: Path
    ) -> bool:
        """
        执行实际的转换操作
        
        Returns:
            是否转换成功
        """
        cv = None
        try:
            cv = Converter(str(pdf_path))
            
            # 如果启用调试模式，生成调试文件
            if enable_debug or self.config['debug']['enable']:
                debug_dir = output_dir / "debug"
                debug_dir.mkdir(exist_ok=True)
                
                # 对第一页生成调试信息（可扩展到所有页）
                try:
                    cv.debug_page(
                        i=0,
                        docx_filename=str(debug_dir / "debug_page_0.docx"),
                        debug_pdf=str(debug_dir / "debug_page_0.pdf"),
                        layout_file=str(debug_dir / "layout_page_0.json"),
                        **kwargs
                    )
                    self.logger.info(f"  调试文件已生成: {debug_dir}")
                except Exception as e:
                    self.logger.warning(f"  生成调试文件失败: {e}")
            
            # 执行转换
            cv.convert(str(docx_path), start=0, end=None, **kwargs)
            
            return True
            
        except Exception as e:
            self.logger.error(f"  转换过程出错: {e}")
            return False
        finally:
            if cv:
                cv.close()
    
    def batch_convert(
        self,
        input_dir: Optional[str] = None,
        output_dir: Optional[str] = None,
        enable_debug: bool = False
    ):
        """
        批量转换目录下的所有 PDF 文件
        
        Args:
            input_dir: 输入目录路径（None 则使用配置文件中的路径）
            output_dir: 输出目录路径（None 则使用配置文件中的路径）
            enable_debug: 是否启用调试模式
        """
        # 确定输入输出目录
        in_dir = Path(input_dir) if input_dir else Path(self.config['input_dir'])
        out_dir = Path(output_dir) if output_dir else Path(self.config['output_dir'])
        
        if not in_dir.exists():
            self.logger.error(f"输入目录不存在: {in_dir}")
            return
        
        # 创建输出目录
        out_dir.mkdir(parents=True, exist_ok=True)
        
        # 查找所有 PDF 文件
        pdf_files = sorted(in_dir.glob("*.pdf"))
        
        if not pdf_files:
            self.logger.warning(f"在 {in_dir} 中未找到 PDF 文件")
            return
        
        self.logger.info(f"找到 {len(pdf_files)} 个 PDF 文件")
        self.logger.info(f"输出目录: {out_dir}")
        self.logger.info("-" * 60)
        
        # 统计信息
        stats = {
            'total': len(pdf_files),
            'success': 0,
            'failed': 0,
            'fallback': 0,
            'total_time': 0
        }
        
        # 逐个转换
        for idx, pdf_path in enumerate(pdf_files, 1):
            print(f"\n[{idx}/{stats['total']}] 正在处理: {pdf_path.name}")
            
            result = self.convert_single(
                pdf_path=pdf_path,
                output_dir=out_dir,
                enable_debug=enable_debug
            )
            
            stats['total_time'] += result['duration']
            
            if result['success']:
                stats['success'] += 1
                if result['use_fallback']:
                    stats['fallback'] += 1
                print(f"  ✓ 成功 ({result['duration']:.2f}s)")
                if result['use_fallback']:
                    print(f"    (使用 fallback 配置)")
            else:
                stats['failed'] += 1
                print(f"  ✗ 失败: {result['message']}")
                if not self.config['error_handling']['continue_on_error']:
                    self.logger.error("遇到错误，停止批量转换")
                    break
        
        # 输出统计信息
        print("\n" + "=" * 60)
        print("转换完成！统计信息：")
        print(f"  总文件数: {stats['total']}")
        print(f"  成功: {stats['success']}")
        print(f"  失败: {stats['failed']}")
        print(f"  使用 fallback: {stats['fallback']}")
        print(f"  总耗时: {stats['total_time']:.2f}s")
        print(f"  平均耗时: {stats['total_time']/stats['total']:.2f}s/文件")
        print("=" * 60)


def main():
    """主函数 - 命令行入口"""
    parser = argparse.ArgumentParser(description='PDF 转 DOCX 批量转换工具')
    parser.add_argument(
        '--input-dir',
        type=str,
        help='输入 PDF 文件目录（默认: config.yaml 中的配置）'
    )
    parser.add_argument(
        '--output-dir',
        type=str,
        help='输出 DOCX 文件目录（默认: config.yaml 中的配置）'
    )
    parser.add_argument(
        '--single',
        type=str,
        help='转换单个 PDF 文件（提供文件路径）'
    )
    parser.add_argument(
        '--config',
        type=str,
        default='config.yaml',
        help='配置文件路径（默认: config.yaml）'
    )
    parser.add_argument(
        '--debug',
        action='store_true',
        help='启用调试模式（生成布局分析文件）'
    )
    
    args = parser.parse_args()
    
    # 初始化转换器
    converter = PDFConverter(config_path=args.config)
    
    # 单文件转换模式
    if args.single:
        pdf_path = Path(args.single)
        if not pdf_path.exists():
            print(f"错误: 文件不存在 - {pdf_path}")
            return
        
        output_dir = Path(args.output_dir) if args.output_dir else Path(converter.config['output_dir'])
        output_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"转换单个文件: {pdf_path.name}")
        result = converter.convert_single(
            pdf_path=pdf_path,
            output_dir=output_dir,
            enable_debug=args.debug
        )
        
        if result['success']:
            print(f"✓ 转换成功!")
            print(f"  输出: {result['output_path']}")
            print(f"  耗时: {result['duration']:.2f}s")
        else:
            print(f"✗ 转换失败: {result['message']}")
    
    # 批量转换模式
    else:
        converter.batch_convert(
            input_dir=args.input_dir,
            output_dir=args.output_dir,
            enable_debug=args.debug
        )


if __name__ == '__main__':
    main()

