function RemoveQuotes(const Str: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(Str) do
    if Str[i] <> '"' then
      Result := Result + Str[i];
end;

procedure ParseCSV(const LineOfText: String; Fields: TStrings);
var
  inQuotes: boolean;
  i: Integer;
  FieldStart: Integer;
begin
  Fields.Clear;
  inQuotes := False;
  FieldStart := 1;
  for i := 1 to Length(LineOfText) do
  begin
    if LineOfText[i] = '"' then
      inQuotes := not inQuotes
    else if (LineOfText[i] = ',') and not inQuotes then
    begin
      Fields.Add(RemoveQuotes(Trim(Copy(LineOfText, FieldStart, i - FieldStart))));
      FieldStart := i + 1;
    end;
  end;
  Fields.Add(RemoveQuotes(Trim(Copy(LineOfText, FieldStart, Length(LineOfText) - FieldStart + 1))));
end;

Procedure AddParamToSchematic;
Var
  CurrentSchematic : ISch_Document;
  Iterator : ISch_Iterator;
  ParamIterator : ISch_Iterator;
  Component : ISch_Component;
  Param : ISch_Parameter;
  ParamExists : Boolean;
  Fields, DesignatorFields, HeaderFields : TStringList;
  i, j : Integer;
  ComponentExists : Boolean;
  CSVFilePath, LineOfText, ParamName, DesignatorColumnName, Report: String;
  ParamColumnIndex, DesignatorColumnIndex : Integer;
  F : TextFile;
Begin
  // Get the current schematic document
  CurrentSchematic := SchServer.GetCurrentSchDocument;
  If CurrentSchematic = Nil Then
  begin
       ShowMessage('No schematic document is open.');
       Exit;
  end;

  //User inputs
  CSVFilePath := InputBox('Enter the CSV file path or name', 'CSV File Path/Name (include .csv)', '');
  ParamName := InputBox('Enter the parameter name to update or add', 'Parameter Name', '');
  DesignatorColumnName := InputBox('Enter the reference designator column name', 'Designator Column Name', '');

  If (CSVFilePath = '') Or (ParamName = '') Or (DesignatorColumnName = '') Then Exit;
  If Not FileExists(CSVFilePath) Then
  Begin
      ShowMessage('File does not exist');
      Exit;
  End;

  Fields := TStringList.Create;
  HeaderFields := TStringList.Create;
  DesignatorFields := TStringList.Create;
  Report := '';

  try
    // Open the file
    AssignFile(F, CSVFilePath);
    Reset(F);

    // Read the header line
    if not EOF(F) then
    begin
      ReadLn(F, LineOfText);
      ParseCSV(LineOfText, HeaderFields);
      ParamColumnIndex := HeaderFields.IndexOf(ParamName);
      DesignatorColumnIndex := HeaderFields.IndexOf(DesignatorColumnName);

      if ParamColumnIndex = -1 then
      begin
        ShowMessage('Parameter name not found in CSV header');
        Exit;
      end;

      if DesignatorColumnIndex = -1 then
      begin
        ShowMessage('Designator column name not found in CSV header');
        Exit;
      end;
    end;

    // Read the rest of the file
    while Not EOF(F) do
    begin
      ReadLn(F, LineOfText);
      ParseCSV(LineOfText, Fields);

      // Check if this line has enough fields
      if Fields.Count >= 2 then
      begin
        // Designator is in the DesignatorColumnIndex field, parameter value is in the ParamColumnIndex field
        DesignatorFields.Delimiter := ',';
        DesignatorFields.DelimitedText := Trim(Fields[DesignatorColumnIndex]);
        //ShowMessage('Before trimming: ' + DesignatorFields.Text); //   test function to test the trimming before delimters remove
        for i := 0 to DesignatorFields.Count - 1 do
        DesignatorFields[i] := Trim(DesignatorFields[i]);
        //ShowMessage('After trimming: ' + DesignatorFields.Text); // test function to test the trimming after the delimters are removed


        for i := 0 to DesignatorFields.Count - 1 do
        Begin
          // Create an iterator to find the component
          Iterator := CurrentSchematic.SchIterator_Create;
          Iterator.AddFilter_ObjectSet(MkSet(eSchComponent));
          ComponentExists := False;

          Component := Iterator.FirstSchObject;
          While Component <> Nil Do
          Begin
            // Check if this is the component you want
            If Component.Designator.Text = Trim(DesignatorFields.Strings[i]) Then
            Begin
              ComponentExists := True;

              // Create an iterator to find the parameter
              ParamExists := False;
              ParamIterator := Component.SchIterator_Create;
              ParamIterator.AddFilter_ObjectSet(MkSet(eParameter));

              Param := ParamIterator.FirstSchObject;
              While Param <> Nil Do
              Begin
                If Param.Name = ParamName Then
                Begin
                  Param.Text := Fields[ParamColumnIndex];
                  ParamExists := True;
                  // Add to report
                  Report := Report + 'Component ' + Trim(DesignatorFields.Strings[i]) + ' parameter ' + ParamName + ' updated.' + sLineBreak;
                  Break;
                End;
                Param := ParamIterator.NextSchObject;
              End;

              Component.SchIterator_Destroy(ParamIterator);

              // If the parameter doesn't exist, create a new one
              If Not ParamExists Then
              Begin
                   Param := SchServer.SchObjectFactory(eParameter, eCreate_Default);
                   Param.Name := ParamName;
                   Param.Text := Fields[ParamColumnIndex];

                   // Position the parameter at the top middle of the component
                   with Component.BoundingRectangle do
                   begin
                        Param.Location := Point((x1 + x2) div 2, y1);
                   end;

                   // Add the new parameter to the component
                   Component.AddSchObject(Param);

                   // Add to report
                   Report := Report + 'Component ' + Trim(DesignatorFields.Strings[i]) + ' parameter ' + ParamName + ' added.' + sLineBreak;
              End;

              // Refresh the document to reflect changes
              CurrentSchematic.GraphicallyInvalidate;
              Break;
            End;

            Component := Iterator.NextSchObject;
          End;

          // Clean up the iterator
          CurrentSchematic.SchIterator_Destroy(Iterator);

          // If the component doesn't exist, add it to the report
          If Not ComponentExists Then
          Begin
            Report := Report + 'Component ' + Trim(DesignatorFields.Strings[i]) + ' does not exist.' + sLineBreak;
          End;
        End;
      End;
    end;

    // After reading the entire file, if there were any updates or missing components, show a message
    if Report <> '' then
      ShowMessage(Report);

  finally
    CloseFile(F);
    Fields.Free;
    HeaderFields.Free;
    DesignatorFields.Free;
  end;
End;

