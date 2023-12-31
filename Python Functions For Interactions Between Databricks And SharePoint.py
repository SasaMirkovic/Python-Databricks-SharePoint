#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# FUNCTION FOR DOWNLOADING FROM SHAREPOINT TO DBFS

def download_files_to_dbfs(site_url, file_url, dbfs_path):
    try:
        credentials = ClientCredential("*****", "*****")
        ctx = ClientContext(site_url).with_credentials(credentials)
        
        list_source = ctx.web.get_folder_by_server_relative_url(file_url)
        files = list_source.files
        ctx.load(files)
        ctx.execute_query()
        
        for myfiles in files:
            rel_url = myfiles.properties["ServerRelativeUrl"]
            download_path = dbfs_path + myfiles.properties["Name"]
            
            with open(download_path, "wb") as local_file:
                file = ctx.web.get_file_by_server_relative_path(rel_url).download(local_file).execute_query()
                print("Downloaded file " + myfiles.properties["Name"])
                
    except Exception as e:
        print(e)


# In[ ]:


# FUNCTION FOR CREATING SHAREPOINT FOLDER

def create_sharepoint_folder (ctx, relative_url, folder_name):
  parent_folder = ctx.web.get_folder_by_server_relative_url(relative_url)
  new_folder = parent_folder.folders.add(folder_name)
  ctx.load(new_folder)
  ctx.execute_query()

  return new_folder


# In[ ]:


# FUNCTION FOR UPLOADING TO SHAREPOINT THAT USES PREVIOUSLY CREATED FUNCTION FOR FOLDER CREATION

def upload_to_sharepoint(site_url, relative_url, dataframe, folder_name, file_name):

    try:
        credentials = ClientCredential("*****", 
                                       "*****")
        ctx = ClientContext(site_url).with_credentials(credentials)

        # Get the SharePoint folder to upload the file to
        try:
          folder = ctx.web.get_folder_by_server_relative_url(relative_url + "/" + folder_name)
          ctx.load(folder)
          ctx.execute_query()

        except:
          folder = create_sharepoint_folder(ctx, relative_url, folder_name)

        # Check if the dataframe is already a Pandas dataframe
        if isinstance(dataframe, pd.DataFrame):
          pandas_df = dataframe

        else:
          # Convert the Spark DataFrame to a pandas DataFrame and then to a CSV string
          pandas_df = dataframe.toPandas()

        excel_bytes = io.BytesIO()

        with pd.ExcelWriter(excel_bytes, engine = 'openpyxl', mode = 'xlsx', if_sheet_exists = 'replace') as writer:
          pandas_df.to_excel(writer, index=False)

        excel_bytes.seek(0)

        # Upload the CSV file to SharePoint
        uploaded_file = folder.upload_file(file_name, excel_bytes).execute_query()
        print("Uploaded file " + uploaded_file.properties["Name"])

    except Exception as e:
        print(e)



# In[ ]:

# FUNCTION FOR UPLOADING TO SHAREPOINT WHICH TAKES A GIVEN FILE NAME TO DETERMINE FILE TYPE BASED ON THE FILE EXTENSION

def upload_files_to_sharepoint(site_url, relative_url, dataframe, file_name):
    try:
        credentials = ClientCredential("*******", "*******")
        ctx = ClientContext(site_url).with_credentials(credentials)


        # Get the SharePoint folder to upload the file to
        folder = ctx.web.get_folder_by_server_relative_url(relative_url)

        # Check if the dataframe is already a Pandas dataframe
        if isinstance(dataframe, pd.DataFrame):
          pandas_df = dataframe
        else:
          # Convert the Spark DataFrame to a pandas DataFrame and then to a CSV string
          pandas_df = dataframe.toPandas()

          # Check file extension
        upload_type = file_name.split('.')[-1]

          # For PARQUET file
        if upload_type == 'parquet':
          # Save the Pandas DataFrame as a Parquet file in memory
          parquet_bytes = io.BytesIO()
          table = pa.Table.from_pandas(pandas_df)
          pq.write_table(table, parquet_bytes)

          parquet_bytes.seek(0)
 
          # Upload the PARQUET file to SharePoint
          uploaded_file = folder.upload_file(file_name, parquet_bytes).execute_query()
          print("Uploaded file " + uploaded_file.properties["Name"])

           # For CSV file 
        if upload_type == 'csv':
          # Save Pandas DataFrame as CSV file in memory
          csv_string = pandas_df.to_csv(index = False)
          csv_bytes = io.BytesIO(csv_string.encode())

           # Upload the CSV file to SharePoint
          uploaded_file = folder.upload_file(file_name, csv_bytes).execute_query()
          print("Uploaded file " + uploaded_file.properties["Name"])

            # For XLSX file
        if upload_type == 'xlsx':
            # Save Pandas DataFrame as XLSX file in memory
          excel_bytes = io.BytesIO()
          with pd.ExcelWriter(excel_bytes, engine = 'openpyxl', mode = 'xlsx', if_sheet_exists = 'replace') as writer:
            pandas_df.to_excel(writer, index=False)

          excel_bytes.seek(0)
            
          # Upload the excel file to SharePoint
          uploaded_file = folder.upload_file(file_name, excel_bytes).execute_query()
          print("Uploaded file " + uploaded_file.properties["Name"])

    except Exception as e:
        print(e)


# In[ ]:


# FUNCTION FOR REMOVING STRING "NaN", "nan" AND "-" VALUES

def replace_nan(df):
  # Iterate over each column in the dataframe
  for column in df.columns:

    # Replace 'NaN' or 'nan' strings with None
    df[column] = df[column].replace(['NaN', 'nan', '-', 'NaT'], [None, None, None, None])

  return df


# In[ ]:


# ITERATING THROUGH FOLDER FOR READING EXCEL FILES

li = []
os.chdir(r'******')
allFiles = glob.glob("*.xlsx")
for file in allFiles :
  df = pd.read_excel(file, sheet_name= '****', engine='openpyxl', skiprows=19, skipfooter = 8, usecols = 'B,F,J:N,P,T')
  li.append(df)
  Excel_files = pd.concat(li)


# In[ ]:


# REMOVING FILES FROM DBFS FOLDER

for i in dbutils.fs.ls("/FileStore/Data/Folder/Subfolder/"):
  dbutils.fs.rm(i[0], True)


# In[ ]:


# # ITERATING THROUGH FOLDER FOR READING EXCEL FILES, WITH AN ADDITIONAL STEP WHICH DROPS ALL ROWS FROM THE FIRST NULL VALUE IN COLUMN A

li = []
os.chdir(r'/dbfs/FileStore/Data/***/****/')
allFiles = glob.glob("*.xlsx")

for file in allFiles :
  df = pd.read_excel(file, engine='openpyxl', skiprows=5)

  # Find the row number of the first null cell in column A
  first_null_row = df[df['Column A'].isnull()].index[0]

  # Calculate the number of rows to skip from the end
  skip_from_end = len(df) - first_null_row

  # Read the file again, skipping the appropriate number of rows from the end
  df = pd.read_excel(file, engine = 'openpyxl', skiprows=5, skipfooter=skip_from_end)
  li.append(df)

  # Concat files into one Dataframe
  DataFrame = pd.concat(li)
