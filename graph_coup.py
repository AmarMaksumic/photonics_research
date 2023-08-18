'''
Write a python script that asks for the name of an excel file, and then uses that file to create a graph.
'''
import os
import psutil
import sys
import time
import matplotlib.pyplot as plt
import pandas as pd
from scipy import signal as sig
from multiprocessing import Process

X_AXIS = 'Wavelength (nm)'
Y_AXIS = 'Transmission (dB)'

def do_the_thing(out_folder, file, in_folders, title, x_axis, y_axis, show_graph):

  plt.title(title)
  plt.xlabel(x_axis)
  plt.ylabel(y_axis)

  mid_line = 0

  for folder in in_folders:
    file_path = folder + '\\' + file
    df = pd.read_excel(file_path, sheet_name="Average IL")
    df.drop(df.tail(1).index,inplace=True) 
    df['Ch.1'] *= -1
    Q1 = df['Ch.1'].quantile(0.25)
    Q3 = df['Ch.1'].quantile(0.75)
    IQR = Q3 + Q1
    df_wo = df[~((df['Ch.1'] < (Q1 - 1.5 * IQR)) |(df['Ch.1'] > (Q3 + 1.5 * IQR)))]
    avg = df['Ch.1'].mean()
    # mid_line += IQR / 2
    print(IQR/2)
    df['Ch.1'] += (-35-avg)
    df.loc[(df['Ch.1'] > 0 )|(df['Ch.1'] < -80 ), 'Ch.1'] = -35
    df['Ch.1'] = sig.savgol_filter(df['Ch.1'], 101, 3)
    plt.plot(df['Wavelength'], df['Ch.1'])

  # mid_line /= len(in_folders)
  # plt.axhline(y = mid_line, color = 'r', linestyle = '-')

  legend_drawn_flag = True
  plt.legend(in_folders, loc=0, frameon=legend_drawn_flag)
  # Save the graph
  file = title
  plt.savefig(out_folder + '\\' + file + '.png', dpi=1000)

  # Show the graph
  if show_graph == 'y':
    plt.show()

  # Clear the graph
  plt.clf()

def get_process_memory():
  process = psutil.Process(os.getpid())
  return process.memory_info().rss

def main():
  start_time = time.time()
  mem_before = get_process_memory()
  # Get command line argument to toggle showing or not showing the graph
  if len(sys.argv) > 1:
    show_graph = sys.argv[1]
  else:
    show_graph = 'y'

  # Ask for the folders we are reading from in comma separated format, and split them into a list
  folders = input('Enter folder pairs: ')
  folders = folders.split()

  if folders[1] == '==>':
    new_folders = []
    for i in range(0, int(folders[2])):
      new_folders.append('C' + str(i+1) + '.' + folders[0][3])

    print(new_folders)
    folders = new_folders


  proc = []
  folder = folders[0]
  # Create the folder for the pictures
  if not os.path.exists('graphs_combined'):
    os.mkdir('graphs_combined')
  out_folder = 'graphs_combined' 
  # Iterate through all the excel files in the folder
  for file in os.listdir(folder):
    if file.endswith(".xlsx"):
      file_path = folder + '\\' + file
      title = file.split('.xlsx')[0]
      p = Process(target=do_the_thing, args=(out_folder, file, folders, title, X_AXIS, Y_AXIS, show_graph, ))
      proc.append((p, file, title, X_AXIS, Y_AXIS, show_graph))

  for p in proc:
    print('Starting process ' + str(p[0]) 
          + ' for \"' + str(p[1]) 
          + '\" with title \"' + str(p[2]) 
          + '\" and x axis \"' + str(p[3])
          + '\" and y axis \"' + str(p[4])
          + '\" and show graph set to \"' + str(p[5]) + '\"\n')
    p[0].start()

  for p in proc:
    p[0].join()

  print('Finished in ' + str(time.time() - start_time) + ' seconds, with memory usage of ' + str(get_process_memory() - mem_before) + ' bytes.')


if __name__ == '__main__':
  main()




# # Iterate through all the excel files in the folder
# for file in os.listdir(folder):
#   if file.endswith(".xlsx"):
#     file_name = folder + '\\' + file
#     # Read the excel file
#     df = pd.read_excel(file_name, sheet_name="Average IL")
#     df.drop(df.tail(1).index,inplace=True) 
#     df['Ch.1'] *= -1

#     # Create a graph
#     plt.title(file)
#     plt.xlabel('Wavelength (nm)')
#     plt.ylabel('T (dB)')
#     plt.plot(df['Wavelength'], df['Ch.1'])

#     # Save the graph
#     file = file.split('.xlsx')[0]
#     plt.savefig(folder + '_pics' + '\\' + file + '.png')

#     # Show the graph
#     if show_graph == 'y':
#       plt.show()

#     # Clear the graph
#     plt.clf()

