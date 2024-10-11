def main():
    opc = int(input("Controlar planilhas por GPT (1) ou manualmente (2)? "))
    if opc == 1:
        try:
            from GPTxlsx import GPTxlsx_main
            GPTxlsx_main()
        except Exception as error:
            input(error)
    else:
        try:
            from MANUALxlsx import MANUALxlsx_main
            MANUALxlsx_main()
        except Exception as error:
            input(error)

if __name__ == "__main__":
    main()