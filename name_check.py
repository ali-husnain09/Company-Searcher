class name_validations:
    def __checkValid__(self, name):
        self.name = str(name)
        self.valid_name = False
        self.valid_chars = []
        # append companies name into list
        with open("companies.txt", "r") as f:
            for line in f:
                self.valid_chars.append(line.strip())
        # to check name validity
        for index in self.valid_chars:
            if index in (self.name.lower()).split(" "):
                self.valid_name = True
                # returns true if its a valid company name
                return self.valid_name
        else:
            # returns false if its not a valid company name
            return False


if __name__ == "__main__":
    validation_obj = name_validations()
    validation_obj.__checkValid__()
