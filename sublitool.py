import json
import csv
import win32com.client


class SubliTool:

    def __init__(self, lineup='lineup.csv', destination="D:/Programming Workspace/Python Projects/subli-tool/export", config='jersey'):
        # Photoshop
        self.ps_app = win32com.client.Dispatch("Photoshop.Application")
        self.current_doc = self.ps_app.ActiveDocument

        # Application vars
        self.csv_lineup = lineup
        self.destination_folder = destination
        self.selected_config = config

        # Config files
        self.configs = {
            'jersey': 'config/jersey.json',
            'short': 'config/short.json'
        }

    def parse_json(self, json_file):
        with open(json_file) as f:
            result = json.load(f)

        return result

    def parse_csv(self, csv_file):
        result = []

        with open(csv_file) as f:
            rows = csv.DictReader(f)

            for row in rows:
                result.append(row)

        return result

    def clean_columns(self, config_columns):
        result = []

        for column in config_columns:
            item = column.split('#')

            for element in item:
                if element not in result:
                    result.append(element.strip())

        return result

    def check_layers(self, layers):
        counter = 0
        missing = 0

        for element in layers:
            try:
                self.current_doc.ArtLayers[element]
                counter += 1
            except:
                missing += 1

        if len(layers) == counter:
            return True
        else:
            return f'Layers count missing: {missing}'

    def save_to_tiff(self, filename):
        tiff_filename = f"{self.destination_folder}/{filename}.tiff"
        tiff_save_options = win32com.client.Dispatch(
            "Photoshop.TiffSaveOptions")
        tiff_save_options.ImageCompression = 3
        tiff_save_options.EmbedColorProfile = True
        tiff_save_options.LayerCompression = 2
        tiff_save_options.Transparency = True
        tiff_save_options.Layers = True
        self.current_doc.SaveAs(tiff_filename, tiff_save_options,
                                True, 3)

    def resize_jersey(self, width, height):
        self.ps_app.Preferences.RulerUnits = 2

        self.current_doc.ResizeImage(width, height, 150)

    def run(self):
        config = self.parse_json(self.configs['jersey'])
        config_layer_names = list(config['layers'].keys())
        config_layers = config['layers']
        config_columns = list(config['layers'].values())
        config_sizes = config['sizes']
        config_size_allowance = config['size_allowance']

        lineup = self.parse_csv(self.csv_lineup)
        lineup_headers = list(self.parse_csv(self.csv_lineup)[0].keys())

        config_columns.append(config['size_column'])
        config_columns.append(config['file_naming'])
        config_columns = self.clean_columns(config_columns)

        if self.check_layers(config_layer_names):
            result = all(
                element in lineup_headers for element in config_columns)

            if result:
                for row in range(len(lineup)):
                    file_format = config['file_naming'].split('#')
                    file_name = ''

                    for key, val in config_layers.items():
                        column = val.split('#')
                        data = ''

                        if len(column) > 0:
                            for item in column:
                                data += f'{lineup[row][item]} '
                        else:
                            data = lineup[row][column[0]]

                        layer = self.current_doc.ArtLayers[key]
                        layer.TextItem.Contents = data.upper().strip()

                    for item in file_format:
                        file_name += f"{lineup[row][item]} "

                    size = lineup[row][config['size_column']].lower()
                    width = config_sizes[size]['width'] + \
                        config_size_allowance['width']
                    height = config_sizes[size]['height'] + \
                        config_size_allowance['height']

                    self.resize_jersey(width, height)
                    self.save_to_tiff(file_name.strip().upper())

                return "Automation process successful!"
            else:
                return "Columns in configuration file doesn't match to lineup headers"
        else:
            return self.check_layers(config_layer_names)


if __name__ == '__main__':
    subli_tool = SubliTool()
    print(subli_tool.run())
