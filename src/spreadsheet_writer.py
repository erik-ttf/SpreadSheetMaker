# Python Built-in
import json
from collections import OrderedDict
import sys
import re
import xlsxwriter

# TTF Modules
import ttfcore.common.environ as env
import ttfcore.shotgun.base as shotgun_base
from shotgun_api.shotgun_api3.shotgun import Shotgun
from ttfcore.ui.base import BaseWidgetWindow, launch

# TTF Site Packages
from Qt import QtWidgets, QtGui, QtCore


class SpreadSheetData(BaseWidgetWindow):
    _obj_name = "SpreadsheetUI"
    _window_title = "Submission Writer"

    def __init__(self, input_data):
        super(SpreadSheetData, self).__init__()
        self._input_fields = self.read_input_data(input_data)
        self._sg_base = self.connect_to_sg()
        self._sg_proj = self._sg_base.project
        self._sg_call = self._sg_base.sg

        self._playlists = self.get_playlists()
        if self._playlists:
            self.combo_playlists.addItems([p.get('code') for p in self._playlists])
            self.combo_playlists.setCurrentIndex(0)

        self.connect_signals()

    def connect_signals(self):
        self.btn_write_submission.clicked.connect(self.submission_write)

    def read_input_data(self, input_data):
        with open(input_data) as open_file:
            content = json.load(open_file, object_pairs_hook=OrderedDict)
        return content

    def connect_to_sg(self):
        # Connects to shotgun via the ShotgunBase module
        sg_base = shotgun_base.ShotgunBase()
        if not sg_base:
            print 'Error: Failed to connect to Shotgun server!'
            return

        print "Successfully connected to Shotgun TTF secure server!"
        return sg_base

    def get_playlists(self):
        filters = [['project', 'is', self._sg_proj]]
        fields = ['code', 'versions']
        playlists = self._sg_call.find('Playlist', filters, fields)
        return playlists

    def get_current_playlist(self):
        selected_playlist = self.combo_playlists.currentText()
        for each_playlist in self._playlists:
            if selected_playlist in each_playlist.get('code'):
                return each_playlist

    def collect_version_data(self, curr_playlist):
        version_data = list()
        for each_vers in curr_playlist.get('versions'):
            version_fields = self._input_fields.get('Shotgun').values()
            version_fields.append('sg_path_to_meta_data')
            sg_version = self._sg_call.find_one(
                'Version', [['id', 'is', each_vers.get('id')]], version_fields)
            version_data.append(sg_version)
        return version_data

    def load_shot_metadata(self, vers_entity):
        version_name = vers_entity['code']
        comp_meta_path = vers_entity.get('sg_path_to_meta_data')
        if not comp_meta_path:
            print 'Warning: Failed to find sg_path_to_meta_data field for ' \
                  'version: {}'.format(version_name)
            return

        with open(comp_meta_path, 'r') as f:
            comp_json = json.load(f)
            shot_meta_path = comp_json['comp']['data_paths'][0]

        if not shot_meta_path:
            print 'Warning: Failed to find shot metadata' \
                  'for version: {}'.format(version_name)
            return

        with open(shot_meta_path, 'r') as f:
            shot_json = json.load(f)

        if not shot_json:
            print 'Warning: Failed to read shot metadata' \
                  'for version: {}'.format(version_name)
            return

        return shot_json

    def modify_values_on_request(self, header, data, vers_entity):
        new_data = None

        # TODO: Edit Submission Notes and add from comp metadata?

        if 'edit' in vers_entity.get('code'):
            camera_attrs = ['tilt', 'speed', 'height', 'camera type', 'lens type', 'lens']
            if header.lower() in camera_attrs:
                new_data = 'N/A'

            if header.lower() in 'type':
                new_data = data + ' - Edit'

            if header.lower() in 'submitted for':
                new_data = 'Edit for review'
        else:
            camera_attrs = ['tilt', 'speed', 'height']
            if header.lower() in camera_attrs:
                if data != '':
                    if '->' in data:
                        min_val, max_val = data.split('->')
                        new_data = str(round(float(min_val.strip()), 2)) + ' -> ' + str(round(float(max_val.strip()), 2))
                    else:
                        new_data = str(round(float(data), 2))

            if 'lens' == header.lower():
                if data != '':
                    lens_number = re.sub('[a-zA-Z]+', '', data)
                    new_data = str(round(float(lens_number))) + 'mm'

        if header in 'Shot Number':
            new_data = re.sub('[0-9]+_', '', data, 1)

        return new_data

    def modify_height_attrs(self, vers_entity):
        modified_height_values = list()

        shot_data = self.load_shot_metadata(vers_entity)
        # Gets the range min/max values for height, speed and tilt to round up to 2 decimals
        height_min = str(round(shot_data['maya']['frame_data']['realHeight']['range']['min'], 2))
        height_max = str(round(shot_data['maya']['frame_data']['realHeight']['range']['max'], 2))

        real_height_frames = shot_data['maya']['frame_data']['realHeight']['values']
        start_frame = str(sorted([int(f) for f in real_height_frames.keys()])[0])
        modified_height_values.append(str(round(real_height_frames[start_frame], 2)))
        end_frame = str(sorted([int(f) for f in real_height_frames.keys()])[-1])
        modified_height_values.append(str(round(real_height_frames[end_frame], 2)))

        # Sets Min and Max values of realHeight
        modified_height_values.append(height_min)
        modified_height_values.append(height_max)

        return modified_height_values

    def submission_write(self):
        self.lbl_status_text.setText('')

        curr_playlist = self.get_current_playlist()
        if not curr_playlist.get('versions'):
            print 'No versions found for playlist: {}'.format(curr_playlist.get('code'))
            return

        # Opens a new workbook, adding a worksheet to write the data in
        wb = xlsxwriter.Workbook('C:/Users/epetters/Documents/test.xlsx')
        ws = wb.add_worksheet()

        # Writes the headers in a bolded format
        bold_format = wb.add_format({'bold': True})
        headers = self._input_fields.get('Shotgun').keys()
        if 'External' in self._input_fields:
            headers.extend(self._input_fields.get('External').keys())
        for col, head in enumerate(headers):
            ws.write(0, col, head, bold_format)

        version_data = self.collect_version_data(curr_playlist)
        col_values = self._input_fields.get('Shotgun').items()
        row = 1
        col = 0
        for index, each_vers in enumerate(version_data):
            for (header, data) in col_values:
                curr_data = version_data[index].get(data, '')
                if isinstance(curr_data, dict):
                    curr_data = curr_data.get('name')
                if curr_data is None:
                    curr_data = ''
                modified_data = self.modify_values_on_request(header, curr_data, each_vers)
                if modified_data:
                    curr_data = modified_data
                ws.write(row, col, curr_data)
                col += 1
            col = 0
            row += 1

        row = 1
        col = len(col_values)
        for index, each_vers in enumerate(version_data):
            if 'External' in self._input_fields:
                height_attrs = self.modify_height_attrs(each_vers)
                for each_attr in height_attrs:
                    ws.write(row, col, each_attr)
                    col += 1
            col = len(col_values)
            row += 1

        self.lbl_status_text.setText('Submission document saved!')

        wb.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print 'No input data provided for submission writing!'
        sys.exit(0)

    launch(SpreadSheetData, sys.argv[1])