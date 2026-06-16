
from datetime import datetime
import os
import logging
from functions.f_navegar_runt import f_navegar_runt_vim  

def main():
    # Carpeta de log dentro del proyecto
    log_folder = os.path.join(os.path.dirname(__file__), 'Data', 'Log')
    os.makedirs(log_folder, exist_ok=True)

    # Configura el logger con nombre 'main'
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file = os.path.join(log_folder, f"app_{timestamp}.log")
    logging.basicConfig(
        filename=log_file, filemode='w',
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    logger = logging.getLogger('main')
    logger.setLevel(logging.INFO)

    logger.info("Se inicializó la automatización.....")
    print("********************************************")
    print("* Automatización. Consultas por numero VIN *")
    print("********************************************")
    print("Se inicializó la automatización.....")

    # Ejecutar flujo VIN
    f_navegar_runt_vim()
    # o: f_navegar_runt()  # usando el wrapper

    print("La automatización ha terminado")
    logger.info("La automatización ha terminado")

if __name__ == "__main__":
    main()
