import requests
from bs4 import BeautifulSoup
import openpyxl
import xlwings as xw
import time,os

def get_movie_details_tmdb(title):
    search_url = f"https://www.themoviedb.org/search?query={title.replace(' ', '%20')}"
    response = requests.get(search_url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find the first result's link
    first_result = soup.find('a', class_='result')
    if first_result:
        movie_link = "https://www.themoviedb.org" + first_result['href']
        
        # Get movie details from the movie page
        movie_response = requests.get(movie_link)
        movie_soup = BeautifulSoup(movie_response.text, 'html.parser')

        movie_title = movie_soup.find('h2').text.strip()
        overview = movie_soup.find('div', class_='overview').text.strip() if movie_soup.find('div', class_='overview') else 'N/A'
        rating_value = movie_soup.find('span', class_='user_score_chart')['data-percent'] if movie_soup.find('span', class_='user_score_chart') else 'N/A'
        
                # Get cast information
        #cast_list = movie_soup.select('div.cast_scroller li.card')
        cast_list = movie_soup.select('ol.people li.card')
        cast = []
        for actor in cast_list:
             actor_name_tag = actor.find('p').find('a')
             actor_name = actor_name_tag.text.strip() if actor_name_tag else 'N/A'
             character = actor.find('p', class_='character').text.strip()
             cast.append({'actor_name': actor_name, 'character': character}) 

        # Find the director and writer names
        people_section = movie_soup.find('ol', class_='people no_image')
        director = 'N/A'
        writer = 'N/A'
        
        if people_section:
            profiles = people_section.find_all('li', class_='profile')
            for profile in profiles:
                role = profile.find('p', class_='character').text.strip()
                name = profile.find('a').text.strip()
                if 'Director' in role:
                    director = name
                if 'Writer' in role:
                    writer = name
        # Get the image URL
        image_tag = movie_soup.find('img', class_='poster')
        image_url = image_tag['src'] if image_tag else 'N/A'                    

        return {
            'Title': movie_title,
            'Rating': rating_value,
            'Overview': overview,
            'Cast': cast,
            'Director': director,
            'Writer': writer,
            'URL': movie_link,
            'Image URL':image_url
        }
    else:
        return None

# Function to read movie titles from Excel and update the details in the Excel file
def update_excel_with_movie_details(excel_file_path, sheet_name):
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook[sheet_name]
    count=2
    DoEntry=False
    # Iterate over each row in column A
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1):
        cell = row[0]
        movie_title = cell.value
        ColumnB=sheet["B"+str(count)].value
        if ColumnB == None:
            DoEntry=True
            time.sleep(3)  # This will pause the program for 3 seconds
            movie_details = get_movie_details_tmdb(movie_title)
        
            if movie_details:
                # Update the Excel sheet with the movie details
                sheet.cell(row=cell.row, column=2, value=movie_details['Title'].replace('\n', ' '))
                sheet.cell(row=cell.row, column=3, value=movie_details['Overview'])
                sheet.cell(row=cell.row, column=4, value=movie_details['Rating'])
                sheet.cell(row=cell.row, column=5, value=movie_details['Director'])
                sheet.cell(row=cell.row, column=6, value=movie_details['Writer'])
                sheet.cell(row=cell.row, column=7, value=movie_details['URL'])
                
                cast_text = "; ".join([f"{member['actor_name']} as {member['character']}" for member in movie_details['Cast']])
                sheet.cell(row=cell.row, column=8, value=cast_text)

                sheet.cell(row=cell.row, column=9, value=movie_details['Image URL'])

                if movie_details:
                    print(f"Title: {movie_details['Title']}")
                    print(f"Overview: {movie_details['Overview']}")
                    print(f"Rating: {movie_details['Rating']}")
                    print(f"Director: {movie_details['Director']}")
                    print(f"Writer: {movie_details['Writer']}")
                    print(f"URL: {movie_details['URL']}")
                    print(f"Image URL: {movie_details['Image URL']}")

                    print("\nCharacters:")
                    for member in movie_details['Cast']:
                        print(f" - {member['actor_name']} / {member['character']}")
        else:
            print("Skiped Movie "+str(count-1))                        
        count=count+1
        
    if DoEntry:
    # Save the updated Excel file
        
        # Use xlwings to add the VBA macro
        wb = xw.Book(excel_file_path)
        vba_code = r'''
            Sub AddButtons()
                Dim ws As Worksheet
                Set ws = ThisWorkbook.Sheets("Sheet1")
                Dim lastRow As Long
                lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

                Dim btn As Button
                Dim i As Long
                For i = 2 To lastRow
                    Set btn = ws.Buttons.Add(ws.Cells(i, 10).Left, ws.Cells(i, 9).Top, 100, 20)
                    btn.Name = "btnShowImage" & i
                    btn.OnAction = "ShowImage"
                    ' Set the button caption to the value in Column A
                    btn.Caption = ws.Cells(i, 1).Value
                Next i
            End Sub

        Sub ShowImage()
            ' In the Tools References dialog in Developer mode,
            ' scroll down and check the box for
            ' Microsoft Visual Basic for Applications Extensibility 5.3
            ' & Microsoft Forms 2.0 Object Library by inserting a form
            Dim btnName As String
            btnName = Application.Caller
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets("Sheet1")
            Dim rowIndex As Long
            rowIndex = ws.Shapes(btnName).TopLeftCell.Row

            Dim imageURL As String
            imageURL = ws.Cells(rowIndex, 9).Value

            If imageURL <> "N/A" Then
                ' Download the image from the URL
                Dim localFilePath As String
                localFilePath = DownloadImage(imageURL)
                
                If localFilePath <> "" Then
                    ' Create the UserForm dynamically
                    Dim VBComp As VBComponent
                    Set VBComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
                    With VBComp
                        .Properties("Width") = 320
                        .Properties("Height") = 390
                        .Properties("Caption") = ws.Cells(rowIndex, 2).Value
                    End With
                    
                    ' Add an Image control to the UserForm
                    Dim ImgControl As MSForms.Image
                    Set ImgControl = VBComp.Designer.Controls.Add("Forms.Image.1")
                    With ImgControl
                        .Left = 10
                        .Top = 10
                        .Width = 320
                        .Height = 350
                        .Picture = LoadPicture(localFilePath)
                    End With
                    
                    ' Show the dynamically created UserForm and delete it afterwards
                    With VBA.UserForms.Add(VBComp.Name)
                        .Show
                        ' Once the form is closed, remove it from the workbook
                        ThisWorkbook.VBProject.VBComponents.Remove VBComp
                    End With
                Else
                    MsgBox "Failed to download the image.", vbExclamation
                End If
            Else
                MsgBox "No image available for this movie.", vbExclamation
            End If
        End Sub

        Function DownloadImage(ByVal url As String) As String
            Dim httpReq As Object
            Set httpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
            httpReq.Open "GET", url, False
            httpReq.Send

            If httpReq.Status = 200 Then
                Dim stream As Object
                Set stream = CreateObject("ADODB.Stream")
                stream.Open
                stream.Type = 1 ' Binary
                stream.Write httpReq.responseBody
                stream.SaveToFile Environ("TEMP") & "\downloaded_image.jpg", 2 ' Overwrite if exists
                stream.Close
                DownloadImage = Environ("TEMP") & "\downloaded_image.jpg"
            Else
                DownloadImage = ""
            End If
        End Function
            '''
            
        wb.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vba_code)
        wb.save(excel_file_path.replace('.xlsx', '.xlsm'))
        wb.close()

# Example usage
excel_file_path = "My Movie Library.xlsx"
#excel_file_path = "F:/Richard/My Movie Library.xlsx"
if not os.path.isabs(excel_file_path):
   excel_file_path=os.path.join(os.path.expanduser("~"), "Documents")+"\\"+excel_file_path
sheet_name = 'Sheet1'
update_excel_with_movie_details(excel_file_path, sheet_name)


