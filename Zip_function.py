from pathlib import Path
import zipfile

def zip_function (diretorio,file_name):
    zip_config = { # mapeia o nome do arquivo zip com os arquivos que ele terá
    file_name: 'teste[1-3].pdf'
    }

    p = Path(diretorio)
    for zip_name, file_pattern in zip_config.items():
        with zipfile.ZipFile(zip_name, "w") as oZip:
             for f in p.glob(file_pattern):
                oZip.write(f, f.name)
    return(print('Arquivo Zipado'))
            

