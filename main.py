import argparse
from core.orchestrator import run_plant

def main():
    parser = argparse.ArgumentParser(description="RPA SAP")
    parser.add_argument(
        "--plant",
        type=str,
        required=True,
        help="Planta para rodar o RPA ('01-Anchieta' / '02-Taubate' / '03-Curitiba' / '04-SaoCarlos')"
    )
    
    args = parser.parse_args()
    
    print("Iniciando RPA...")
    print("Ctrl+C p fechar.")
    
    try:
        run_plant(args.plant)
    except KeyboardInterrupt:
        print("\nRPA fechado.")

if __name__ == "__main__":
    main()
