import codecs
import os

def foundinfile(openfile, term):
    contents = openfile.readlines()
    found = False

    for line in contents:
        if term.upper() in line.upper():
            found = True

    return found

def main():
    directory = "Enter path to exported files"
    object_name = "Enter object name"
    found = []

    for root, _, files in os.walk(directory):
        for f in files:
            path = os.path.join(root, f)

            # Open without decoding first.
            with open(path, "r") as openfile:
                if foundinfile(openfile, object_name):
                    found.append(f)

            # If nothing found then try decoding.
            if f not in found:
                with codecs.open(path, "r", "utf-16-le") as openfile:
                    if foundinfile(openfile, object_name):
                        found.append(f)

    print("Results for: {0}".format(object_name))
    for f in found:
        print(f)

if __name__ == "__main__":
    main()