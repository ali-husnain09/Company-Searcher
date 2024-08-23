import json


class address_validations:

    def __checkValid__(self, address, city, state):
        self.data = {}
        self.temp = {}
        self.temp[address] = {"city": city, "state": state}
        # Load JSON data from file
        try:
            with open("address_data.json", "r") as infile:  #  if the file is available
                self.data = json.load(infile)
        except json.JSONDecodeError:  # if the file is not available
            with open("address_data.json", "w") as outfile:
                json.dump({}, outfile)
        except FileNotFoundError:
            with open("address_data.json", "w") as outfile:
                json.dump({}, outfile)
        for key in self.temp.keys():
            if key in self.data.keys():
                self.data.clear()
                self.data.update(self.temp)
                # Save the data to a JSON file
                with open("address_data.json", "w") as outfile:
                    json.dump(self.data, outfile)
                return True
            else:
                self.data.clear()
                self.data.update(self.temp)
                # Save the data to a JSON file
                with open("address_data.json", "w") as outfile:
                    json.dump(self.data, outfile)
                return False


if __name__ == "__main__":
    checking_obj = address_validations()
    checking_obj.__checkValid__()
