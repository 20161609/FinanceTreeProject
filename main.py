import shellController as Sc

if __name__ == '__main__':
    shell = Sc.Shell()

    while True:
        command = input(shell.prompt)

        if len(command.split()):
            shell.fetch(command)
            if command in Sc.quit_commands:
                print("Exit..")
                break
