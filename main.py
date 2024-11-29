def main():
    opc = int(input("Controlar planilhas por GPT (1) ou manualmente (2)? "))
    if opc == 1:
        try:
            import GPTxlsx
            GPTxlsx.GPTmain()
        except Exception as error:
            input(error)
    else:
        try:
            import MANUALxlsx
            MANUALxlsx.MANUALxlsx_main()
        except Exception as error:
            input(error)

if __name__ == "__main__":
    main()

#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