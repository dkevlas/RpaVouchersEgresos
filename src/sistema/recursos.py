import subprocess

def recursos():
    try:
        cpu = subprocess.check_output(
            "wmic cpu get loadpercentage /value",
            shell=True,
            encoding="latin-1",
            errors="ignore"
        )

        cpu = cpu.split("=")[-1].strip()

        #mem = subprocess.check_output(cmd, shell=True, encoding="latin-1")
        
        print(f"\nCPU: {cpu.strip()}%")
        #print(f"RAM: {mem.strip()}%\n")

    except Exception as e:
        print(f"Error al traer los recursos: {e}")


if __name__ == "__main__":
    recursos()