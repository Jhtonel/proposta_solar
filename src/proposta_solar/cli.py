#!/usr/bin/env python3
import argparse
import sys
from pathlib import Path
from proposta_solar.presentation import main as gerar_apresentacao

def parse_args():
    parser = argparse.ArgumentParser(
        description='Gerador de Propostas Solares',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  %(prog)s -e input/dados.xlsx -t templates/modelo.pptx -o output/proposta.pptx
  %(prog)s --excel input/dados.xlsx --template templates/modelo.pptx --output output/proposta.pptx
        """
    )
    
    parser.add_argument(
        '-e', '--excel',
        required=True,
        help='Arquivo Excel com os dados'
    )
    
    parser.add_argument(
        '-t', '--template',
        required=True,
        help='Template PowerPoint'
    )
    
    parser.add_argument(
        '-o', '--output',
        required=True,
        help='Caminho para salvar a apresentação'
    )
    
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Exibir logs detalhados'
    )
    
    return parser.parse_args()

def main():
    args = parse_args()
    
    # Converter caminhos para Path
    excel_path = Path(args.excel)
    template_path = Path(args.template)
    output_path = Path(args.output)
    
    # Verificar se os diretórios existem
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    template_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Chamar o gerador
    return gerar_apresentacao(
        excel_path=excel_path,
        template_path=template_path,
        output_path=output_path,
        verbose=args.verbose
    )

if __name__ == '__main__':
    sys.exit(main()) 