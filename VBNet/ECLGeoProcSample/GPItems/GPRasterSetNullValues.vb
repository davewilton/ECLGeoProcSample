Imports System
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesRaster
Imports ESRI.ArcGIS.Geometry

Namespace GPItems

    Public Class GPRasterSetNullValues : Inherits ECLGPBaseFunction

        Friend Overrides ReadOnly Property ToolName() As String
            Get
                Return "RasterSetNull"
            End Get
        End Property

        Friend Overrides ReadOnly Property ToolDisplayName() As String
            Get
                Return "Raster set null values"
            End Get
        End Property

        Public Overrides ReadOnly Property ParameterInfo() As IArray
            Get
                'Array to the hold the parameters
                Dim pParameters As IArray = New ArrayClass()

                Dim inputParameter As IGPParameterEdit3 = New GPParameterClass()

                'The input raster
                inputParameter.DataType = New GPRasterLayerTypeClass()
                inputParameter.Value = New GPRasterLayerClass()
                inputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput
                inputParameter.DisplayName = "Input Raster"
                inputParameter.Name = "in_raster"
                inputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired
                pParameters.Add(inputParameter)

                'operator
                Dim gpStringValue As IGPString
                inputParameter = New GPParameterClass()
                inputParameter.DataType = New GPStringTypeClass()
                'Set the default value
                gpStringValue = New GPStringClass()
                gpStringValue.Value = "equalTo"
                inputParameter.Value = CType(gpStringValue, IGPValue)
                inputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput
                inputParameter.DisplayName = "Operator"
                inputParameter.Name = "in_operator"
                inputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired
                ' Create a fixed list of string values
                Dim cvDomain As IGPCodedValueDomain = New GPCodedValueDomainClass()
                cvDomain.AddStringCode("equalTo", "equalTo")
                cvDomain.AddStringCode("lessThan", "lessThan")
                cvDomain.AddStringCode("greaterThan", "greaterThan")
                inputParameter.Domain = CType(cvDomain, IGPDomain)
                pParameters.Add(inputParameter)

                'operator
                inputParameter = New GPParameterClass()
                'Set the default value
                inputParameter.DataType = CType(New GPDoubleTypeClass(), IGPDataType)
                inputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput
                inputParameter.DisplayName = "Value"
                inputParameter.Name = "in_value"
                inputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired
                pParameters.Add(inputParameter)

                'The out raster
                Dim outputParameter As IGPParameterEdit3 = New GPParameterClass()
                outputParameter.DataType = CType(New DERasterDatasetType(), IGPDataType)
                outputParameter.Value = New GPRasterLayerClass()
                outputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput
                outputParameter.DisplayName = "Output Raster"
                outputParameter.Name = "out_raster"
                outputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired
                pParameters.Add(outputParameter)

                Return pParameters
            End Get
        End Property

        Public Overrides Sub Execute(ByVal paramvalues As IArray, ByVal trackcancel As ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages)
            Try

                Dim parameterValue As IGPValue
                parameterValue = GpUtilities.UnpackGPValue(Parameters.Element(ByName("out_raster")))
                If CType(envMgr, IGeoProcessorSettings).OverwriteOutput Then DeleteExisting(parameterValue)
                Dim outRasterPath As String = parameterValue.GetAsText()

                'get the operators
                Dim doubleNullValue As Double
                Dim opp As String
                parameterValue = GpUtilities.UnpackGPValue(Parameters.Element(ByName("in_value")))
                doubleNullValue = Convert.ToDouble(parameterValue.GetAsText())
                parameterValue = GpUtilities.UnpackGPValue(Parameters.Element(ByName("in_operator")))
                opp = parameterValue.GetAsText()

                'get the input raster
                Dim rasterDS As IRasterDataset2 = Nothing
                parameterValue = GpUtilities.UnpackGPValue(Parameters.Element(ByName("in_raster")))
                GpUtilities.DecodeRasterLayer(parameterValue, rasterDS)

                'Create the new raster
                Dim ws As IRasterWorkspace2
                Dim wsF As IWorkspaceFactory = New RasterWorkspaceFactory
                Dim file As New IO.FileInfo(outRasterPath)
                ws = CType(wsF.OpenFromFile(file.DirectoryName, 0), IRasterWorkspace2)
                Dim raster As IRaster2
                Dim info As IRasterProps
                raster = CType(rasterDS.CreateFullRaster(), IRaster2)
                info = CType(raster, IRasterProps)
                'Create a raster cursor which is one column in width and height
                Dim pnt As IPnt = New PntClass
                pnt.X = info.Width
                pnt.Y = info.Height
                Dim origin As IPoint
                origin = New Point
                origin.PutCoords(info.Extent.XMin, info.Extent.YMin)
                Dim newrasterDS As IRasterDataset2 = CType(ws.CreateRasterDataset(file.Name, "TIFF", origin, info.Width, info.Height, info.MeanCellSize().X, info.MeanCellSize().Y, 1, info.PixelType, info.SpatialReference, True), IRasterDataset2)
                Dim newRaster As IRaster2 = CType(newrasterDS.CreateFullRaster(), IRaster2)
                'Now set the values
                CType(newRaster, IRasterProps).NoDataValue = info.NoDataValue

                Dim bands As IRasterBandCollection
                Dim band As IRasterBand
                bands = CType(newrasterDS, IRasterBandCollection)
                band = bands.Item(0)
                Dim infoNew As IRasterProps = CType(band, IRasterProps)
                infoNew.NoDataValue = info.NoDataValue(0)
                'Create a cursor which is the same size as the raster
                pnt.X = info.Width
                pnt.Y = info.Height

                Dim rasterCursor As IRasterCursor
                rasterCursor = raster.CreateCursorEx(pnt)
                Dim pixelBlock As IPixelBlock3
                pixelBlock = CType(rasterCursor.PixelBlock, IPixelBlock3)
                Dim pixels As System.Array
                pixels = CType(pixelBlock.PixelData(0), System.Array)

                'Create the output cursor (same size)
                Dim rasterCursorNew As IRasterCursor
                rasterCursorNew = newRaster.CreateCursorEx(pnt)
                Dim pixelBlockNew As IPixelBlock3
                pixelBlockNew = CType(rasterCursorNew.PixelBlock, IPixelBlock3)
                Dim pixelsNew As System.Array

                pixelsNew = CType(pixelBlockNew.PixelData(0), System.Array)

                'set the no data value
                'Get the pixel array.

                'Loop through each pixel.
                For i As Integer = 0 To pixelBlock.Width - 1
                    For j As Integer = 0 To pixelBlock.Height - 1
                        Dim val As Object = pixels.GetValue(i, j)

                        'check if the value should be set to null
                        If val = infoNew.NoDataValue Then
                            pixelsNew.SetValue(infoNew.NoDataValue, i, j) 'if it was already null ensure it stays null
                        Else
                            Select Case opp 'otherwise perform the comparison
                                Case "equalTo"
                                    If Math.Abs(Convert.ToDouble(val) - doubleNullValue) < 0.000000001 Then
                                        pixelsNew.SetValue(infoNew.NoDataValue, i, j)
                                    Else
                                        pixelsNew.SetValue(pixels.GetValue(i, j), i, j)
                                    End If
                                Case "lessThan"
                                    If Convert.ToDouble(val) < doubleNullValue Then
                                        pixelsNew.SetValue(infoNew.NoDataValue, i, j)
                                    Else
                                        pixelsNew.SetValue(pixels.GetValue(i, j), i, j)
                                    End If
                                Case "greaterThan"
                                    If Convert.ToDouble(val) > doubleNullValue Then
                                        pixelsNew.SetValue(infoNew.NoDataValue, i, j)
                                    Else
                                        pixelsNew.SetValue(pixels.GetValue(i, j), i, j)
                                    End If
                            End Select
                        End If
                    Next
                Next

                pixelBlockNew.PixelData(0) = pixelsNew
                'Define the location that the upper left corner of the pixel block is to write.
                Dim upperLeft As IPnt
                upperLeft = New Pnt
                upperLeft.SetCoords(0, 0)

                'Write the pixel block.
                Dim rasterEdit As IRasterEdit
                rasterEdit = CType(newRaster, IRasterEdit)
                rasterEdit.Write(upperLeft, CType(pixelBlockNew, IPixelBlock))

                newrasterDS.PrecalculateStats(0) 'must do this otherwise it won't set the null value

                Runtime.InteropServices.Marshal.ReleaseComObject(rasterEdit)
                Runtime.InteropServices.Marshal.ReleaseComObject(rasterDS)
                Runtime.InteropServices.Marshal.ReleaseComObject(raster)
                Runtime.InteropServices.Marshal.ReleaseComObject(pixelBlock)

            Catch ex As Exception
                message.AddError(0, ex.Message)
            End Try

        End Sub

        Public Overrides Sub UpdateParameters(ByVal paramvalues As IArray, ByVal pEnvMgr As IGPEnvironmentManager)
            Parameters = paramvalues

            'force output to be a tif
            Dim parameterValue As IGPValue
            parameterValue = GpUtilities.UnpackGPValue(Parameters.Element(ByName("out_raster")))
            If Not parameterValue Is Nothing Then
                Dim outGrid As String = parameterValue.GetAsText()
                If Not outGrid.EndsWith(".tif") And outGrid.Trim() <> "" Then
                    parameterValue.SetAsText(outGrid & ".tif")
                End If
            End If
        End Sub

        Public Overrides Sub UpdateMessages(ByVal paramvalues As IArray, ByVal pEnvMgr As IGPEnvironmentManager, ByVal messages As IGPMessages)

            'don't allow user to select geodatabase (just to simplify things for sample)
            Dim parameterValue As IGPValue
            parameterValue = GpUtilities.UnpackGPValue(Parameters.Element(ByName("out_raster")))
            If Not parameterValue Is Nothing Then
                Dim outGrid As String = parameterValue.GetAsText()
                If outGrid.Contains(".gdb") Then
                    messages.ReplaceError(ByName("out_workspace"), 0, "Please select a folder not geodatabase for the output raster")
                End If
            End If
        End Sub

    End Class
End Namespace