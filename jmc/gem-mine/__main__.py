#!/usr/bin/env pipenv-shebang

import sys
import logging
import argparse

from xlrd import open_workbook

# (r,g,b)
chest_color = (255, 153, 0)
gem_color = (255, 0, 255)

def setup_log():
    logging.basicConfig(stream=sys.stdout, level=logging.DEBUG)

def parse_args(args):
    parser = argparse.ArgumentParser(description='Determine the most efficient gem mine route')
    parser.add_argument('--mine', dest='mine', required=True, help='Excel file containing the mine topography')
    parser.add_argument('--count', dest='count', default=28, help='Number of gems to mine')
    return parser.parse_args(args)

"""Start with returning list of x,y coordinates of gems w.r.t. chest"""
def import_mine(mine_file):

    wb = open_workbook(mine_file, formatting_info=True)
    sheet = wb.sheet_by_index(0)

    chest_coords = ()
    gem_coord_list = []

    for col_idx in range(1, sheet.ncols):
        for row_idx in range(1, sheet.nrows):
            cell = sheet.cell(row_idx, col_idx)
            fmt = wb.xf_list[cell.xf_index]
            background_color = wb.colour_map[fmt.background.pattern_colour_index]
            logging.debug('({},{}) = {}'.format(col_idx, row_idx, background_color))

            if background_color == gem_color:
                logging.debug('gem!')
                gem_coord_list.append((col_idx, row_idx))

            if background_color == chest_color:
                logging.debug('chest!')
                chest_coords = (col_idx, row_idx)

    logging.debug('chest is at {}'.format(chest_coords))

    recentered_gem_coord_list = []

    # recenter coordinate grid on chest
    for gem_coords in gem_coord_list:

        # x is simple because excel counts column from left to right
        # like your typical coordinate plane; just gem(x) - chest(x)
        new_x = gem_coords[0] - chest_coords[0]
        # y is less intuitive because excel counts rows from top to bottom,
        # while coordinate planes increment from bottom to top; use chest(y) - gem(y)
        new_y = chest_coords[1] - gem_coords[1]

        recentered_gem_coords = (new_x, new_y)
        recentered_gem_coord_list.append(recentered_gem_coords)

    return recentered_gem_coord_list

def main(args):
    args = parse_args(args)
    setup_log()

    mine = import_mine(args.mine)

    logging.debug('final print!')

    logging.debug('there are {} gems'.format(len(mine)))
    for gem in mine:
        logging.debug('gem is located at ({},{}) w.r.t. chest'.format(gem[0], gem[1]))

if __name__ == '__main__':
    main(sys.argv[1:])
