import os
import email
from email import policy
from rich.console import Console
from rich.progress import Progress
import sys


def extract_attachments(eml_folder, output_folder, progress_callback=None):
    console = Console()
    os.makedirs(output_folder, exist_ok=True)
    eml_files = [f for f in os.listdir(eml_folder) if f.lower().endswith(".eml")]
    total = len(eml_files)

    def nombre_corto(nombre, maxlen=30):
        nombre = os.path.basename(nombre)
        return (nombre[:maxlen] + "...") if len(nombre) > maxlen else nombre

    def texto_estado(idx, total, nombre_eml, ancho=60):
        nombre_eml = nombre_corto(nombre_eml)
        texto = f"[cyan]Procesando correos... {idx}/{total} | {nombre_eml}"
        return texto.ljust(ancho)

    def save_attachment(part, output_folder):
        filename = part.get_filename()
        if filename:
            full_path = os.path.join(output_folder, filename)
            os.makedirs(os.path.dirname(full_path), exist_ok=True)
            payload = part.get_payload(decode=True)
            if payload is not None:
                with open(full_path, "wb") as af:
                    af.write(payload)

    if progress_callback is None:
        with Progress(transient=True, expand=True) as progress:
            task = progress.add_task(texto_estado(0, total, ""), total=total)
            for idx, fname in enumerate(eml_files):
                with open(os.path.join(eml_folder, fname), "rb") as f:
                    msg = email.message_from_binary_file(f, policy=policy.default)
                    # Recorrer todas las partes, no solo iter_attachments
                    for part in msg.walk():
                        # Solo guardar si tiene filename (es adjunto)
                        if (
                            part.get_content_disposition() == "attachment"
                            and part.get_filename()
                        ):
                            save_attachment(part, output_folder)
                progress.update(
                    task,
                    advance=1,
                    description=texto_estado(idx + 1, total, fname, ancho=60),
                )
    else:
        for idx, fname in enumerate(eml_files):
            with open(os.path.join(eml_folder, fname), "rb") as f:
                msg = email.message_from_binary_file(f, policy=policy.default)
                for part in msg.walk():
                    if (
                        part.get_content_disposition() == "attachment"
                        and part.get_filename()
                    ):
                        save_attachment(part, output_folder)
            progress_callback(idx + 1, total, fname)


if __name__ == "__main__":
    try:
        extract_attachments(
            os.path.join(os.path.dirname(__file__), "eml"),
            os.path.join(os.path.dirname(__file__), "examples"),
        )
    except KeyboardInterrupt:
        print("\nProcesamiento cancelado por el usuario (Ctrl+C).")
        sys.exit(0)
        print("\nProcesamiento cancelado por el usuario (Ctrl+C).")
        sys.exit(0)
