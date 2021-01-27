import os


def generate_empty_text_input_files():
    for i in range(2010, 2021):
        for q in range(1, 5):
            path = os.path.join("Input", "MafiaPlayerList" + str(i) + "Q" + str(q) + ".txt")
            open(path, 'a').close()


if __name__ == "__main__":
    generate_empty_text_input_files()