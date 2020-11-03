# Python Built-in
import json
from collections import OrderedDict
import sys

# TTF Modules
import ttfcore.common.environ as env
import ttfcore.shotgun.base as shotgun_base
from shotgun_api.shotgun_api3.shotgun import Shotgun
from ttfcore.ui.base import BaseWidgetWindow, launch

# TTF Site Packages
from Qt import QtWidgets, QtGui, QtCore

import xlsxwriter


# TODO: Init the class storing the UI widgets and logic to pull playlists 


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
            sg_version = self._sg_call.find_one(
                'Version', [['id', 'is', each_vers.get('id')]], self._input_fields.get('Shotgun').values())
            version_data.append(sg_version)
        return version_data

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

        row = 1
        for each_vers in range(len(version_data)):
            for col, data in enumerate(self._input_fields.get('Shotgun').values()):
                curr_data = version_data[each_vers].get(data, '')
                if isinstance(curr_data, dict):
                    curr_data = curr_data.get('name')
                ws.write(row, col, curr_data)
            row += 1

        self.lbl_status_text.setText('Submission document saved!')

        wb.close()

def update_submission_versions(shotgun):
    """
    Updates the submission columns for all versions in the project
    by using the realHeight frame data from JSON metadata
    :param shotgun: The ShotgunBase object to query Shotgun information
    :return:
    """
    sg = shotgun.sg
    sg_proj = shotgun.project

    # Finds the shots on the current Shotgun project
    shot_filters = [['project', 'is', sg_proj]]
    shot_fields = ['id', 'code', 'tasks']

    all_shots = sg.find('Shot', shot_filters, shot_fields)

    # Filters shots under the Previs task
    previs_shots = filter_shots_by_task(shots=all_shots, task_name='Previs')
    if not previs_shots:
        print 'Error: Issues to find previs shots in project: {}'.format(sg_proj.get('name'))
        return

    for each_shot in previs_shots:
        print "\nSearching for versions in shot: {}...".format(each_shot.get('code'))
        version_filters = [['project', 'is', sg_proj], ['entity', 'is', each_shot]]
        version_fields = ['sg_path_to_meta_data', 'id', 'code', 'sg_external_update']

        versions = sg.find('Version', version_filters, version_fields)

        if not versions:
            print 'Error: Issues to find versions for shot with code: {}' \
                .format(each_shot.get('code'))
            continue

        new_versions = [v for v in versions if v.get('sg_external_update') is None \
                            or v.get('sg_external_update') != 'Yes']
        if not new_versions:
            print 'Versions already updated for shot with code: {}' \
                .format(each_shot.get('code'))
            continue

        for each_vers in new_versions:
            version_name = each_vers['code']
            comp_meta_path = each_vers.get('sg_path_to_meta_data')
            if not comp_meta_path:
                print 'Warning: Failed to find sg_path_to_meta_data field for ' \
                      'version: {}'.format(version_name)
                continue

            with open(comp_meta_path, 'r') as f:
                comp_json = json.load(f)
                shot_meta_path = comp_json['comp']['data_paths'][0]

            if not shot_meta_path:
                print 'Warning: Failed to find shot metadata' \
                      'for version: {}'.format(version_name)
                continue

            shot_json = None
            with open(shot_meta_path, 'r') as f:
                shot_json = json.load(f)

            if not shot_json:
                print 'Warning: Failed to read shot metadata' \
                      'for version: {}'.format(version_name)
                continue

            print "Updating version: {}...".format(version_name)

            # Creates a dict to store values for updating version fields in Shotgun
            submission_fields = dict()

            # Gets the range min/max values for height, speed and tilt to round up to 2 decimals
            height_min = str(round(shot_json['maya']['frame_data']['realHeight']['range']['min'], 2))
            height_max = str(round(shot_json['maya']['frame_data']['realHeight']['range']['max'], 2))

            speed_min = str(round(shot_json['maya']['frame_data']['speed']['range']['min'], 2))
            speed_max = str(round(shot_json['maya']['frame_data']['speed']['range']['max'], 2))

            tilt_min = str(round(shot_json['maya']['frame_data']['realTilt']['range']['min'], 2))
            tilt_max = str(round(shot_json['maya']['frame_data']['realTilt']['range']['max'], 2))

            submission_fields['sg_height'] = '{min} -> {max}'.format(min=height_min, max=height_max)
            submission_fields['sg_speed'] = '{min} -> {max}'.format(min=speed_min, max=speed_max)
            submission_fields['sg_tilt'] = '{min} -> {max}'.format(min=tilt_min, max=tilt_max)

            tilt_min = str(round(shot_json['maya']['frame_data']['realTilt']['range']['min'], 2))
            tilt_max = str(round(shot_json['maya']['frame_data']['realTilt']['range']['max'], 2))

            # Gets min/max for lens to append 'mm' unit to the values
            lens_min = shot_json['maya']['frame_data']['Lens']['range']['min']
            lens_max = shot_json['maya']['frame_data']['Lens']['range']['max']
            if lens_min == lens_max:
            	new_lens = str(lens_min) + 'mm'
            else:
        		new_lens = '{min}mm -> {max}mm'.format(min=lens_min, max=lens_max)

            submission_fields['sg_lens_1'] = new_lens

            # Gets all frame values for realHeight
            real_height_frames = shot_json['maya']['frame_data']['realHeight']['values']

            # Sets the first and start frame values for realHeight
            start_frame = str(sorted([int(f) for f in real_height_frames.keys()])[0])
            submission_fields['sg_height_frame_start'] = str(round(real_height_frames[start_frame], 2))
            end_frame = str(sorted([int(f) for f in real_height_frames.keys()])[-1])
            submission_fields['sg_height_end_frame'] = str(round(real_height_frames[end_frame], 2))

            # Sets Min and Max values of realHeight
            submission_fields['sg_height_true_min'] = height_min
            submission_fields['sg_height_true_max'] = height_max 

            # Sets the condition field to 'Yes' to skip already processed versions next time script is run  
            submission_fields['sg_external_update'] = 'Yes'

            # Updating height fields in Shotgun for current version
            sg.update(entity_type='Version', entity_id=each_vers['id'], data=submission_fields)

    print '\nFinished updating submission columns for versions in Shotgun project: {}'.format(sg_proj.get('name'))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print 'No input data provided for submission writing!'
        sys.exit(0)

    launch(SpreadSheetData, sys.argv[1])