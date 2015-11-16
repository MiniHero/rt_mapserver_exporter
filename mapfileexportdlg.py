# -*- coding: utf-8 -*-

"""
/***************************************************************************
Name                 : RT MapServer Exporter
Description          : A plugin to export qgs project to mapfile
Date                 : Oct 21, 2012
copyright            : (C) 2012 by Giuseppe Sucameli (Faunalia)
email                : brush.tyler@gmail.com

 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""

from PyQt4.QtCore import *
from PyQt4.QtGui import *

from qgis.core import *
from qgis.gui import *
from PyQt4 import QtGui

from .ui.interface import Ui_MapfileExportDlg
import mapscript
import re
from mapscript import MS_IMAGEMODE_RGB
from osgeo._ogr import Layer_Intersection
from decimal import Decimal
from owslib.csw import outputformat
import SerializationUtils as utils

_toUtf8 = lambda s: unicode(s).encode('utf8')


defaultEnableRequest = "*"
defaultTitle = "QGIS-MAP"
defaultOnlineResource = "http://localhost/cgi-bin/mapserv"
srsList = []
ms_map = mapscript.mapObj()
    
class MapfileExportDlg(QDialog, Ui_MapfileExportDlg):

    unitMap = {
        QGis.DecimalDegrees : mapscript.MS_DD,
        QGis.Meters : mapscript.MS_METERS,
        QGis.Feet : mapscript.MS_FEET
    }

    onOffMap = {
        True : mapscript.MS_ON,
        False : mapscript.MS_OFF
    }

    trueFalseMap = {
        True : mapscript.MS_TRUE,
        False : mapscript.MS_FALSE
    }
    
    trueFalseStringMap = {
        "TRUE" : mapscript.MS_TRUE,
        "FALSE" : mapscript.MS_FALSE
    }
    
    labelTypeMap = {
        "TRUETYPE" : mapscript.MS_TRUETYPE,
        "BITMAP" : mapscript.MS_BITMAP
    }
    
    labelPositionMap = {
        "AUTO" : mapscript.MS_AUTO,
        "UL" : mapscript.MS_UL,
        "UC" : mapscript.MS_UC,
        "UR" : mapscript.MS_UR,
        "CL" : mapscript.MS_CL,
        "CC" : mapscript.MS_CC,
        "CR" : mapscript.MS_CR,
        "LL" : mapscript.MS_LL,
        "LC" : mapscript.MS_LC,
        "LR" : mapscript.MS_LR
    }

    imageTypeCmbMap = {
        "png" : 0,
        "gif" : 1,
        "jpeg" : 2,
        "svg" : 3,
        "GTiff" : 4
    }
    
    outputFormatOptionMap = {
        "RGB" : mapscript.MS_IMAGEMODE_RGB,
        "RGBA" : mapscript.MS_IMAGEMODE_RGBA,
        "BYTE" : mapscript.MS_IMAGEMODE_BYTE,
        "FEATURE" : mapscript.MS_IMAGEMODE_FEATURE,
        "INT16" : mapscript.MS_IMAGEMODE_INT16,
        "NULL" : mapscript.MS_IMAGEMODE_NULL,
        "PC256" : mapscript.MS_IMAGEMODE_PC256,
        "FLOAT32" : mapscript.MS_IMAGEMODE_FLOAT32
    }
    
    outputFormatImageModeIndexMap = {
        "RGB" : 0,
        "RGBA" : 1,
        "BYTE" : 2,
        "FEATURE" : 3,
        "INT16" : 4,
        "NULL" : 5,
        "PC256" : 6,
        "FLOAT32" : 7
    }
    
    PROJ_LIB = "PROJ_LIB"
    MS_ERRORFILE = "MS_ERRORFILE"
    MS_DEBUGLEVEL = "MS_DEBUGLEVEL"
    
    @classmethod
    def getLayerType(self, layer):
        if layer.type() == QgsMapLayer.RasterLayer:
            return mapscript.MS_LAYER_RASTER
        if layer.geometryType() == QGis.Point:
            return mapscript.MS_LAYER_POINT
        if layer.geometryType() == QGis.Line:
            return mapscript.MS_LAYER_LINE
        if layer.geometryType() == QGis.Polygon:
            return mapscript.MS_LAYER_POLYGON

    @classmethod
    def getLabelPosition(self, palLabel):
        quadrantPosition = palLabel.quadOffset  
        if quadrantPosition == QgsPalLayerSettings.QuadrantAboveLeft: # y=1 x=-1 
            return mapscript.MS_UL
        if quadrantPosition == QgsPalLayerSettings.QuadrantAbove: # y=1 x=0
            return mapscript.MS_UC
        if quadrantPosition == QgsPalLayerSettings.QuadrantAboveRight: # y=1 x=1
            return mapscript.MS_UR
        if quadrantPosition == QgsPalLayerSettings.QuadrantLeft: # y=0 x=-1
            return mapscript.MS_CL
        if quadrantPosition == QgsPalLayerSettings.QuadrantOver: # y=0 x=0
            return mapscript.MS_CC
        if quadrantPosition == QgsPalLayerSettings.QuadrantRight: # y=0 x=1
            return mapscript.MS_CR
        if quadrantPosition == QgsPalLayerSettings.QuadrantBelowLeft: # y=-1 x=-1 
            return mapscript.MS_LL
        if quadrantPosition == QgsPalLayerSettings.QuadrantBelow: # y=-1 x=0
            return mapscript.MS_LC
        if quadrantPosition == QgsPalLayerSettings.QuadrantBelowRight: # y=-1 x=1
            return mapscript.MS_LR
        return mapscript.MS_AUTO

    def __init__(self, iface, parent=None):
        QDialog.__init__(self, parent)
        self.setupUi(self)
        self.iface = iface
        self.canvas = self.iface.mapCanvas()
        self.legend = self.iface.legendInterface()

        # hide map unit combo and label
        self.label4.hide()
        self.cmbGeneralMapUnits.hide()
        srsList.append( _toUtf8( self.canvas.mapRenderer().destinationCrs().authid() ) )

        # setup the template table
        m = TemplateModel(self)
        for layer in self.legend.layers():
            m.append( layer )
            
        self.templateTable.setModel(m)
        d = TemplateDelegate(self)
        self.templateTable.setItemDelegate(d)
        
        # get the default title from the project
        title = QgsProject.instance().title()
        if title == "":
            title = QFileInfo( QgsProject.instance().fileName() ).completeBaseName()
        if title != "":
            self.txtGeneralMapName.setText( title )

        # fill the image format combo
        self.cmbGeneralMapImageType.addItems( ["png", "gif", "jpeg", "svg", "GTiff"] )
        
        self.cmbOutputFormatImageMode_1.addItems(["", "RGB", "RGBA", "BYTE", "FEATURE", "INT16", "NULL", "PC256", "FLOAT32"])
        self.cmbOutputFormatImageMode_2.addItems(["", "RGB", "RGBA", "BYTE", "FEATURE", "INT16", "NULL", "PC256", "FLOAT32"])
        self.cmbOutputFormatImageMode_3.addItems(["", "RGB", "RGBA", "BYTE", "FEATURE", "INT16", "NULL", "PC256", "FLOAT32"])
        self.cmbOutputFormatImageMode_4.addItems(["", "RGB", "RGBA", "BYTE", "FEATURE", "INT16", "NULL", "PC256", "FLOAT32"])
        self.cmbOutputFormatTransparent_1.addItems(["", "TRUE", "FALSE"])
        self.cmbOutputFormatTransparent_2.addItems(["", "TRUE", "FALSE"])
        self.cmbOutputFormatTransparent_3.addItems(["", "TRUE", "FALSE"])
        self.cmbOutputFormatTransparent_4.addItems(["", "TRUE", "FALSE"])
        
        self.cmbOutputFormatImageMode_1.setCurrentIndex(1)
        self.cmbOutputFormatImageMode_2.setCurrentIndex(1)
        self.cmbOutputFormatTransparent_1.setCurrentIndex(2)

        QObject.connect( self.btnChooseFile, SIGNAL("clicked()"), self.selectMapFile )
        QObject.connect( self.toolButtonImportMapFile, SIGNAL("clicked()"), self.importFromMapfile )
        QObject.connect( self.btnChooseTemplate, SIGNAL("clicked()"), self.selectTemplateBody )
        QObject.connect( self.btnChooseTmplHeader, SIGNAL("clicked()"), self.selectTemplateHeader )
        QObject.connect( self.btnChooseTmplFooter, SIGNAL("clicked()"), self.selectTemplateFooter )
        
        self.groupBoxMetadataOws.setEnabled(False)
        self.groupBoxMetadataWms.setEnabled(False)
        self.groupBoxMetadataWfs.setEnabled(False)
        self.groupBoxMetadataWcs.setEnabled(False)
        self.checkBoxOws.stateChanged.connect(self.toggleOwsMetadata)
        self.checkBoxWms.stateChanged.connect(self.toggleWmsMetadata)
        self.checkBoxWfs.stateChanged.connect(self.toggleWfsMetadata)
        self.checkBoxWcs.stateChanged.connect(self.toggleWcsMetadata)
        

    def toggleOwsMetadata(self):
        if self.checkBoxOws.isChecked():
            self.groupBoxMetadataOws.setEnabled(True)
            self.txtMetadataOwsOwsEnableRequest.setText(defaultEnableRequest)
            self.txtMetadataOwsGmlFeatureid.setText("gid")
            self.txtMetadataOwsGmlIncludeItems.setText("all")
            self.txtMetadataOwsOwsTitle.setText(defaultTitle)
            self.txtMetadataOwsOwsOnlineResource.setText(defaultOnlineResource)
            self.txtMetadataOwsOwsSrs.setText(' '.join(srsList))
        else:
            self.txtMetadataOwsOwsEnableRequest.setText("")
            self.txtMetadataOwsGmlFeatureid.setText("")
            self.txtMetadataOwsGmlIncludeItems.setText("")
            self.txtMetadataOwsOwsTitle.setText("")
            self.txtMetadataOwsOwsOnlineResource.setText("")
            self.txtMetadataOwsOwsSrs.setText("")
            self.groupBoxMetadataOws.setEnabled(False)
            
    def toggleWmsMetadata(self):
        if self.checkBoxWms.isChecked():
            self.groupBoxMetadataWms.setEnabled(True)
            self.txtMetadataWmsWebWmsEnableRequest.setText(defaultEnableRequest)
            self.txtMetadataWmsWebWmsTitle.setText(defaultTitle)
            self.txtMetadataWmsWebWmsOnlineresource.setText(defaultOnlineResource)
            self.txtMetadataWmsWebWmsSrs.setText(' '.join(srsList))
        else:
            self.txtMetadataWmsWebWmsEnableRequest.setText("")
            self.txtMetadataWmsWebWmsTitle.setText("")
            self.txtMetadataWmsWebWmsOnlineresource.setText("")
            self.txtMetadataWmsWebWmsSrs.setText("")
            self.groupBoxMetadataWms.setEnabled(False)
        
    def toggleWfsMetadata(self):
        if self.checkBoxWfs.isChecked():
            self.groupBoxMetadataWfs.setEnabled(True)
            self.txtMetadataWfsWebWfsEnableRequest.setText(defaultEnableRequest)
            self.txtMetadataWfsWebWfsTitle.setText(defaultTitle)
            self.txtMetadataWfsWebWfsOnlineresource.setText(defaultOnlineResource)
            self.txtMetadataWfsWebWfsSrs.setText(' '.join(srsList))
        else:
            self.txtMetadataWfsWebWfsEnableRequest.setText("")
            self.txtMetadataWfsWebWfsTitle.setText("")
            self.txtMetadataWfsWebWfsOnlineresource.setText("")
            self.txtMetadataWfsWebWfsSrs.setText("")
            self.groupBoxMetadataWfs.setEnabled(False)
            
    def toggleWcsMetadata(self):
        if self.checkBoxWcs.isChecked():
            self.groupBoxMetadataWcs.setEnabled(True)
            self.txtMetadataWcsWebWcsEnableRequest.setText(defaultEnableRequest)
        else:
            self.txtMetadataWcsWebWcsEnableRequest.setText("")
            self.groupBoxMetadataWcs.setEnabled(False)
            
    def selectMapFile(self):
        # retrieve the last used map file path
        settings = QSettings()
        lastUsedFile = settings.value("/rt_mapserver_exporter/lastUsedFile", "", type=str)

        # ask for choosing where to store the map file
        filename = QFileDialog.getSaveFileName(self, "Select where to save the map file", lastUsedFile, "MapFile (*.map)")
        if filename == "":
            return

        # store the last used map file path
        settings.setValue("/rt_mapserver_exporter/lastUsedFile", filename)
        # update the displayd path
        self.txtMapFilePath.setText( filename )
        
    def importFromMapfile(self):
        # retrieve the last used map file path
        settings = QSettings()
        lastUsedFile = settings.value("/rt_mapserver_exporter/lastUsedFile", "", type=str)

        # ask for choosing where to store the map file
        filename = QFileDialog.getSaveFileName(self, "Select where to save the map file", lastUsedFile, "MapFile (*.map)")
        if filename == "":
            return

        # store the last used map file path
        settings.setValue("/rt_mapserver_exporter/lastUsedFile", filename)
        # update the displayd path
        self.txtMapFilePath.setText( filename )
        ms_map = mapscript.mapObj(filename)
        
        self.txtGeneralMapName.setText(ms_map.name)
        self.cmbGeneralMapImageType.setCurrentIndex(self.imageTypeCmbMap.get(ms_map.imagetype))
        self.txtGeneralMapWidth.setText(str(ms_map.get_width()))
        self.txtGeneralMapHeight.setText(str(ms_map.get_height()))
        self.txtGeneralMapProjLibFolder.setText(ms_map.getConfigOption(self.PROJ_LIB))
        self.txtMetadataMapMsErrorFilePath.setText(ms_map.getConfigOption(self.MS_ERRORFILE))
        self.txtMetadataMapDebugLevel.setText(ms_map.getConfigOption(self.MS_DEBUGLEVEL))
        if ms_map.web.metadata.get("ows_onlineresource") is not None:
            self.txtGeneralWebServerUrl.setText(ms_map.web.metadata.get("ows_onlineresource").split('?')[0])
        self.txtGeneralWebImagePath.setText(ms_map.web.imagepath)
        self.txtGeneralWebImageUrl.setText(ms_map.web.imageurl)
        self.txtGeneralWebTempPath.setText(ms_map.web.temppath)
        self.txtGeneralExternalGraphicRegexp.setText(ms_map.web.validation.get("sld_external_graphic"))
        self.txtMapFontsetPath.setText(ms_map.fontset.filename)
        
        self.txtOutputFormatName_1.setText(ms_map.outputformat.name)
        self.txtOutputFormatDriver_1.setText(ms_map.outputformat.driver)
        self.txtOutputFormatMimetype_1.setText(ms_map.outputformat.mimetype)
        self.txtOutputFormatExtension_1.setText(ms_map.outputformat.extension)
        self.txtOutputFormatImageMode_1.setText(ms_map.outputformat.imagemode)
        self.txtOutputFormatTransparent_1.setText(ms_map.outputformat.transparent)
        self.txtOutputFormatFormatOption_1_1.setText(ms_map.outputformat.formatoptions)
#         self.txtOutputFormatName_1.setText(ms_map.outputformat.name)
    
        
        self.txtMetadataOwsOwsEnableRequest.setText(ms_map.web.metadata.get("ows_enable_request"))
        self.txtMetadataOwsOwsTitle.setText(ms_map.web.metadata.get("ows_title"))
        self.txtMetadataOwsOwsSrs.setText(ms_map.web.metadata.get("ows_srs"))
        if ms_map.web.metadata.get("ows_onlineresource") is not None:
            self.txtMetadataOwsOwsOnlineResource.setText(ms_map.web.metadata.get("ows_onlineresource").split('?')[0])
        self.txtMetadataOwsWebOwsAllowedIpList.setText(ms_map.web.metadata.get("ows_allowed_ip_list"))
        self.txtMetadataOwsWebOwsDeniedIpList.setText(ms_map.web.metadata.get("ows_denied_ip_list"))
        self.txtMetadataOwsWebOwsSchemasLocation.setText(ms_map.web.metadata.get("ows_schemas_location"))
        self.txtMetadataOwsWebOwsUpdatesequence.setText(ms_map.web.metadata.get("ows_updatesequence"))
        self.txtMetadataOwsWebOwsHttpMaxAge.setText(ms_map.web.metadata.get("ows_http_max_age"))
        self.txtMetadataOwsWebOwsSldEnabled.setText(ms_map.web.metadata.get("ows_sld_enabled"))
        self.txtMetadataOwsOwsGetFeatureFormatList.setText(ms_map.web.metadata.get("ows_getfeature_formatlist"))
         
        self.txtMetadataWmsWebWmsEnableRequest.setText(ms_map.web.metadata.get("wms_enable_request"))
        self.txtMetadataWmsWebWmsTitle.setText(ms_map.web.metadata.get("wms_title"))
        if ms_map.web.metadata.get("wms_onlineresource") is not None:
            self.txtMetadataWmsWebWmsOnlineresource.setText(ms_map.web.metadata.get("wms_onlineresource").split('?')[0])
        self.txtMetadataWmsWebWmsSrs.setText(ms_map.web.metadata.get("wms_srs"))
        self.txtMetadataWmsWebWmsAttrbutionOnlineresource.setText(ms_map.web.metadata.get("wms_attribution_onlineresource"))
        self.txtMetadataWmsWebWmsAttributionTitle.setText(ms_map.web.metadata.get("wms_attribution_title"))
        self.txtMetadataWmsWebWmsBboxExtended.setText(ms_map.web.metadata.get("wms_bbox_extended"))
        self.txtMetadataWmsWebWmsFeatureInfoMimeType.setText(ms_map.web.metadata.get("wms_feature_info_mime_type"))
        self.txtMetadataWmsWebWmsEncoding.setText(ms_map.web.metadata.get("wms_encoding"))
        self.txtMetadataWmsWebWmsGetcapabilitiesVersion.setText(ms_map.web.metadata.get("wms_getcapabilities_version"))
        self.txtMetadataWmsWebWmsFees.setText(ms_map.web.metadata.get("wms_fees"))
        self.txtMetadataWmsWebWmsGetmapFormallist.setText(ms_map.web.metadata.get("wms_getmap_formatlist"))
        self.txtMetadataWmsWebWmsGetlegendgraphicFormailist.setText(ms_map.web.metadata.get("wms_getlegendgraphic_formatlist"))
        self.txtMetadataWmsWebWmsKeywordlistVocabulary.setText(ms_map.web.metadata.get("wms_keywordlist_vocabulary"))
        self.txtMetadataWmsWebWmsKeywordlist.setText(ms_map.web.metadata.get("wms_keywordlist"))
        self.txtMetadataWmsWebWmsLanguages.setText(ms_map.web.metadata.get("wms_languages"))
        self.txtMetadataWmsWebWmsLayerlimit.setText(ms_map.web.metadata.get("wms_layerlimit"))
        self.txtMetadataWmsWebWmsRootlayerKeywordlist.setText(ms_map.web.metadata.get("wms_rootlayer_keywordlist"))
        self.txtMetadataWmsWebWmsRemoteSldMaxBytes.setText(ms_map.web.metadata.get("wms_remote_sld_max_bytes"))
        self.txtMetadataWmsWebWmsServiceOnlineResource.setText(ms_map.web.metadata.get("wms_service_onlineresource"))
        self.txtMetadataWmsWebWmsResx.setText(ms_map.web.metadata.get("wms_resx"))
        self.txtMetadataWmsWebWmsResy.setText(ms_map.web.metadata.get("wms_resy"))
        self.txtMetadataWmsWebWmsRootlayerAbstract.setText(ms_map.web.metadata.get("wms_rootlayer_abstract"))
        self.txtMetadataWmsWebWmsRootlayerTitle.setText(ms_map.web.metadata.get("wms_rootlayer_title"))
        self.txtMetadataWmsWebWmsTimeformat.setText(ms_map.web.metadata.get("wms_timeformat"))
        
        self.txtMetadataWfsWebWfsEnableRequest.setText(ms_map.web.metadata.get("wfs_enable_request"))
        self.txtMetadataWfsWebWfsTitle.setText(ms_map.web.metadata.get("wfs_title"))
        if ms_map.web.metadata.get("wfs_onlineresource") is not None:
            self.txtMetadataWfsWebWfsOnlineresource.setText(ms_map.web.metadata.get("wfs_onlineresource").split('?')[0])
        self.txtMetadataWfsWebWfsAbstract.setText(ms_map.web.metadata.get("wfs_abstract"))
        self.txtMetadataWfsWebWfsAccessconstraints.setText(ms_map.web.metadata.get("wfs_accessconstraints"))
        self.txtMetadataWfsWebWfsEncoding.setText(ms_map.web.metadata.get("wfs_encoding"))
        self.txtMetadataWfsWebWfsFeatureCollection.setText(ms_map.web.metadata.get("wfs_feature_collection"))
        self.txtMetadataWfsWebWfsFees.setText(ms_map.web.metadata.get("wfs_fees"))
        self.txtMetadataWfsWebWfsKeywordlist.setText(ms_map.web.metadata.get("wfs_keywordlist"))
        self.txtMetadataWfsWebWfsGetcapabilitiesVersion.setText(ms_map.web.metadata.get("wfs_getcapabilities_version"))
        self.txtMetadataWfsWebWfsNamespacePrefix.setText(ms_map.web.metadata.get("wfs_namespace_prefix"))
        self.txtMetadataWfsWebWfsMaxfeatures.setText(ms_map.web.metadata.get("wfs_maxfeatures"))
        self.txtMetadataWfsWebWfsServiceOnlineresource.setText(ms_map.web.metadata.get("wfs_service_onlineresource"))
        self.txtMetadataWfsWebWfsNamespaceUri.setText(ms_map.web.metadata.get("wfs_namespace_uri"))
        self.txtMetadataOwsOwsGetFeatureFormatList.setText(ms_map.web.metadata.get("wfs_getfeature_formatlist"))
        
        self.txtMetadataWcsWebWcsEnableRequest.setText(ms_map.web.metadata.get("wcs_enable_request"))
        self.txtMetadataWcsWebWcsLabel.setText(ms_map.web.metadata.get("wcs_label"))
        self.txtMetadataWcsWebWcsAbstract.setText(ms_map.web.metadata.get("wcs_abstract"))
        self.txtMetadataWcsWebWcsAccessconstraints.setText(ms_map.web.metadata.get("wcs_accessconstraints"))
        self.txtMetadataWcsWebWcsDescription.setText(ms_map.web.metadata.get("wcs_description"))
        self.txtMetadataWcsWebWcsKeywords.setText(ms_map.web.metadata.get("wcs_keywords"))
        self.txtMetadataWcsWebWcsFees.setText(ms_map.web.metadata.get("wcs_fees"))
        self.txtMetadataWcsWebWcsMetadatalinkFormat.setText(ms_map.web.metadata.get("wcs_metadatalink_format"))
        self.txtMetadataWcsWebWcsMetadatalinkHref.setText(ms_map.web.metadata.get("wcs_metadatalink_href"))
        self.txtMetadataWcsWebWcsMetadatalinkType.setText(ms_map.web.metadata.get("wcs_metadatalink_type"))
        self.txtMetadataWcsWebWcsName.setText(ms_map.web.metadata.get("wcs_name"))
        self.txtMetadataWcsWebWcsServiceOnlineresource.setText(ms_map.web.metadata.get("wcs_service_onlineresource"))
        
        self.txtTemplatePath.setText(ms_map.web.template)
        self.txtTmplHeaderPath.setText(ms_map.web.header)
        self.txtTmplFooterPath.setText(ms_map.web.footer)
        

    def selectTemplateBody(self):
        self.selectTemplateFile( self.txtTemplatePath )

    def selectTemplateHeader(self):
        self.selectTemplateFile( self.txtTmplHeaderPath )

    def selectTemplateFooter(self):
        self.selectTemplateFile( self.txtTmplFooterPath )

    def selectTemplateFile(self, lineedit):
        # retrieve the last used template file path
        settings = QSettings()
        lastUsedFile = settings.value("/rt_mapserver_exporter/lastUsedTmpl", "", type=str)

        # ask for choosing where to store the map file
        filename = QFileDialog.getOpenFileName(self, "Select the template file", lastUsedFile, "Template (*.html *.tmpl);;All files (*);;")
        if filename == "":
            return

        # store the last used map file path
        settings.setValue("/rt_mapserver_exporter/lastUsedTmpl", filename)
        # update the path
        lineedit.setText( filename )

    def accept(self):
        # check user inputs
        if self.txtMapFilePath.text() == "":
            QMessageBox.warning(self, "RT MapServer Exporter", "Mapfile output path is required")
            return

        # create a new ms_map
#         ms_map = mapscript.mapObj()
        ms_map.name = _toUtf8( self.txtGeneralMapName.text() )

        # map size
        width, height = int(self.txtGeneralMapWidth.text()), int(self.txtGeneralMapHeight.text())
        widthOk, heightOk = isinstance(width, int), isinstance(height, int)
        if widthOk and heightOk:
            ms_map.setSize( width, height )

        # map units
        ms_map.units = self.unitMap[ self.canvas.mapUnits() ]
        if self.cmbGeneralMapUnits.currentIndex() >= 0:
            units, ok = self.cmbGeneralMapUnits.itemData( self.cmbGeneralMapUnits.currentIndex() )
            if ok:
                ms_map.units = units

        # map extent
        extent = self.canvas.fullExtent()
        ms_map.extent.minx = extent.xMinimum()
        ms_map.extent.miny = extent.yMinimum()
        ms_map.extent.maxx = extent.xMaximum()
        ms_map.extent.maxy = extent.yMaximum()
        ms_map.setProjection( _toUtf8( self.canvas.mapRenderer().destinationCrs().toProj4() ) )
        
        if self.txtGeneralMapProjLibFolder.text().strip(' \t\n\r') != "":
            ms_map.setConfigOption(self.PROJ_LIB, self.txtGeneralMapProjLibFolder.text().strip(' \t\n\r'))
        if self.txtMetadataMapMsErrorFilePath.text().strip(' \t\n\r') != "":
            ms_map.setConfigOption(self.MS_ERRORFILE, self.txtMetadataMapMsErrorFilePath.text().strip(' \t\n\r'))
        if self.txtMetadataMapDebugLevel.text().strip(' \t\n\r') != "":
            ms_map.setConfigOption(self.MS_DEBUGLEVEL, self.txtMetadataMapDebugLevel.text().strip(' \t\n\r'))

#         if self.txtGeneralMapShapePath.text() != "":
#             ms_map.shapepath = _toUtf8( self.txtGeneralMapShapePath.text() )

        # image section
        r,g,b,a = self.canvas.canvasColor().getRgb()
        ms_map.imagecolor.setRGB( r, g, b )    #255,255,255
        ms_map.setImageType( _toUtf8( self.cmbGeneralMapImageType.currentText() ) )
        ms_outformat = ms_map.getOutputFormatByName( ms_map.imagetype )
#         ms_outformat.transparent = self.onOffMap[ True ]

        # legend section
        #r,g,b,a = self.canvas.canvasColor().getRgb()
        #ms_map.legend.imageColor.setRgb( r, g, b )
        #ms_map.legend.status = mapscript.MS_ON
        #ms_map.legend.keysizex = 18
        #ms_map.legend.keysizey = 12
        #ms_map.legend.label.type = mapscript.MS_BITMAP
        #ms_map.legend.label.size = MEDIUM??
        #ms_map.legend.label.color.setRgb( 0, 0, 89 )
#         ms_map.legend.label.partials = self.trueFalseMap[ self.checkBoxPartials ]
#         ms_map.legend.label.force = self.trueFalseMap[ self.checkBoxForce ]
        #ms_map.legend.template = "[templatepath]"

        # web section
        ms_map.web.imagepath = _toUtf8( self.txtGeneralWebImagePath.text() )
        ms_map.web.imageurl = _toUtf8( self.txtGeneralWebImageUrl.text() )
        if self.txtGeneralWebTempPath.text() != "":
            ms_map.web.temppath = _toUtf8( self.txtGeneralWebTempPath.text() )
        # add validation block if set a regexp
        # no control on regexp => it will be done by mapscript applySld
        # generating error in case regexp is wrong
        validationRegexp = _toUtf8( self.txtGeneralExternalGraphicRegexp.text() )
        if validationRegexp != "":
            ms_map.web.validation.set("sld_external_graphic", validationRegexp)

        # web template
        ms_map.web.template = _toUtf8( self.getTemplatePath() )
        
        if self.txtTmplHeaderPath.text() != "":
            ms_map.web.header = _toUtf8( self.txtTmplHeaderPath.text() )
            
        if self.txtTmplFooterPath.text() != "":
            ms_map.web.footer = _toUtf8( self.txtTmplFooterPath.text() )
            
        if self.txtGeneralMapMaxSize.text() != "":
            ms_map.maxsize = int(self.txtGeneralMapMaxSize.text())
        
        # outputformat node
        if self.txtOutputFormatName_1.text() != "":
#             outputformat1 = mapscript.outputFormatObj('AGG/PNG', 'myPNG')
            outputformat1 = ms_map.getOutputFormatByName(_toUtf8( self.cmbGeneralMapImageType.currentText() ))
            outputformat1.inmapfile = mapscript.MS_TRUE
            outputformat1.name = self.txtOutputFormatName_1.text()
            if self.txtOutputFormatDriver_1.text() != "":
                outputformat1.driver = self.txtOutputFormatDriver_1.text()
            if self.txtOutputFormatMimetype_1.text() != "":
                outputformat1.mimetype = self.txtOutputFormatMimetype_1.text()
            if self.txtOutputFormatExtension_1.text() != "":
                outputformat1.extension = self.txtOutputFormatExtension_1.text()
            if self.cmbOutputFormatImageMode_1.currentText() != "":
                outputformat1.imagemode = self.outputFormatOptionMap.get(_toUtf8(self.cmbOutputFormatImageMode_1.currentText()))
            if self.cmbOutputFormatTransparent_1.currentText() != "":
                outputformat1.transparent = self.trueFalseStringMap.get(_toUtf8(self.cmbOutputFormatTransparent_1.currentText()))
            if self.txtOutputFormatFormatOptionKey_1_1.text() != "" and self.txtOutputFormatFormatOptionValue_1_1.text() != "":
                outputformat1.setOption(str(self.txtOutputFormatFormatOptionKey_1_1.text()), str(self.txtOutputFormatFormatOptionValue_1_1.text()))
            if self.txtOutputFormatFormatOptionKey_2_1.text() != "" and self.txtOutputFormatFormatOptionValue_2_1.text() != "":
                outputformat1.setOption(str(self.txtOutputFormatFormatOptionKey_2_1.text()), str(self.txtOutputFormatFormatOptionValue_2_1.text()))
            if self.txtOutputFormatFormatOptionKey_3_1.text() != "" and self.txtOutputFormatFormatOptionValue_3_1.text() != "":
                outputformat1.setOption(str(self.txtOutputFormatFormatOptionKey_3_1.text()), str(self.txtOutputFormatFormatOptionValue_3_1.text()))
            if self.txtOutputFormatFormatOptionKey_4_1.text() != "" and self.txtOutputFormatFormatOptionValue_4_1.text() != "":
                outputformat1.setOption(str(self.txtOutputFormatFormatOptionKey_4_1.text()), str(self.txtOutputFormatFormatOptionValue_4_1.text()))
            ms_map.appendOutputFormat(outputformat1)
        
        if self.txtOutputFormatName_2.text() != "":
            outputformat2 = mapscript.outputFormatObj('AGG/PNG', self.txtOutputFormatName_2.text())
#             outputformat1 = ms_map.getOutputFormatByName(_toUtf8( self.cmbGeneralMapImageType.currentText() ))
            outputformat2.inmapfile = mapscript.MS_TRUE
            outputformat2.name = self.txtOutputFormatName_2.text()
            if self.txtOutputFormatDriver_2.text() != "":
                outputformat2.driver = self.txtOutputFormatDriver_2.text()
            if self.txtOutputFormatMimetype_2.text() != "":
                outputformat2.mimetype = self.txtOutputFormatMimetype_2.text()
            if self.txtOutputFormatExtension_2.text() != "":
                outputformat2.extension = self.txtOutputFormatExtension_2.text()
            if self.cmbOutputFormatImageMode_2.currentText() != "":
                outputformat2.imagemode = self.outputFormatOptionMap.get(_toUtf8(self.cmbOutputFormatImageMode_2.currentText()))
            if self.cmbOutputFormatTransparent_2.currentText() != "":
                outputformat2.transparent = self.trueFalseStringMap.get(_toUtf8(self.cmbOutputFormatTransparent_2.currentText()))
            if self.txtOutputFormatFormatOptionKey_1_2.text() != "" and self.txtOutputFormatFormatOptionValue_1_2.text() != "":
                outputformat2.setOption(str(self.txtOutputFormatFormatOptionKey_1_2.text()), str(self.txtOutputFormatFormatOptionValue_1_2.text()))
            if self.txtOutputFormatFormatOptionKey_2_2.text() != "" and self.txtOutputFormatFormatOptionValue_2_2.text() != "":
                outputformat2.setOption(str(self.txtOutputFormatFormatOptionKey_2_2.text()), str(self.txtOutputFormatFormatOptionValue_2_2.text()))
            if self.txtOutputFormatFormatOptionKey_3_2.text() != "" and self.txtOutputFormatFormatOptionValue_3_2.text() != "":
                outputformat2.setOption(str(self.txtOutputFormatFormatOptionKey_3_2.text()), str(self.txtOutputFormatFormatOptionValue_3_2.text()))
            if self.txtOutputFormatFormatOptionKey_4_2.text() != "" and self.txtOutputFormatFormatOptionValue_4_2.text() != "":
                outputformat2.setOption(str(self.txtOutputFormatFormatOptionKey_4_2.text()), str(self.txtOutputFormatFormatOptionValue_4_2.text()))
            ms_map.appendOutputFormat(outputformat2)
        
        if self.txtOutputFormatName_3.text() != "":
            outputformat3 = mapscript.outputFormatObj('AGG/PNG', self.txtOutputFormatName_3.text())
#             outputformat1 = ms_map.getOutputFormatByName(_toUtf8( self.cmbGeneralMapImageType.currentText() ))
            outputformat3.inmapfile = mapscript.MS_TRUE
            outputformat3.name = self.txtOutputFormatName_3.text()
            if self.txtOutputFormatDriver_3.text() != "":
                outputformat3.driver = self.txtOutputFormatDriver_3.text()
            if self.txtOutputFormatMimetype_3.text() != "":
                outputformat3.mimetype = self.txtOutputFormatMimetype_3.text()
            if self.txtOutputFormatExtension_3.text() != "":
                outputformat3.extension = self.txtOutputFormatExtension_3.text()
            if self.cmbOutputFormatImageMode_3.currentText() != "":
                outputformat3.imagemode = self.outputFormatOptionMap.get(_toUtf8(self.cmbOutputFormatImageMode_3.currentText()))
            if self.cmbOutputFormatTransparent_3.currentText() != "":
                outputformat3.transparent = self.trueFalseStringMap.get(_toUtf8(self.cmbOutputFormatTransparent_3.currentText()))
            if self.txtOutputFormatFormatOptionKey_1_3.text() != "" and self.txtOutputFormatFormatOptionValue_1_3.text() != "":
                outputformat3.setOption(str(self.txtOutputFormatFormatOptionKey_1_3.text()), str(self.txtOutputFormatFormatOptionValue_1_3.text()))
            if self.txtOutputFormatFormatOptionKey_2_3.text() != "" and self.txtOutputFormatFormatOptionValue_2_3.text() != "":
                outputformat3.setOption(str(self.txtOutputFormatFormatOptionKey_2_3.text()), str(self.txtOutputFormatFormatOptionValue_2_3.text()))
            if self.txtOutputFormatFormatOptionKey_3_3.text() != "" and self.txtOutputFormatFormatOptionValue_3_3.text() != "":
                outputformat3.setOption(str(self.txtOutputFormatFormatOptionKey_3_3.text()), str(self.txtOutputFormatFormatOptionValue_3_3.text()))
            if self.txtOutputFormatFormatOptionKey_4_3.text() != "" and self.txtOutputFormatFormatOptionValue_4_3.text() != "":
                outputformat3.setOption(str(self.txtOutputFormatFormatOptionKey_4_3.text()), str(self.txtOutputFormatFormatOptionValue_4_3.text()))
            ms_map.appendOutputFormat(outputformat3)
            
        if self.txtOutputFormatName_4.text() != "":
            outputformat4 = mapscript.outputFormatObj('AGG/PNG', self.txtOutputFormatName_4.text())
#             outputformat1 = ms_map.getOutputFormatByName(_toUtf8( self.cmbGeneralMapImageType.currentText() ))
            outputformat4.inmapfile = mapscript.MS_TRUE
            outputformat4.name = self.txtOutputFormatName_4.text()
            if self.txtOutputFormatDriver_4.text() != "":
                outputformat4.driver = self.txtOutputFormatDriver_4.text()
            if self.txtOutputFormatMimetype_4.text() != "":
                outputformat4.mimetype = self.txtOutputFormatMimetype_4.text()
            if self.txtOutputFormatExtension_4.text() != "":
                outputformat4.extension = self.txtOutputFormatExtension_4.text()
            if self.cmbOutputFormatImageMode_4.currentText() != "":
                outputformat4.imagemode = self.outputFormatOptionMap.get(_toUtf8(self.cmbOutputFormatImageMode_4.currentText()))
            if self.cmbOutputFormatTransparent_4.currentText() != "":
                outputformat4.transparent = self.trueFalseStringMap.get(_toUtf8(self.cmbOutputFormatTransparent_4.currentText()))
            if self.txtOutputFormatFormatOptionKey_1_4.text() != "" and self.txtOutputFormatFormatOptionValue_1_4.text() != "":
                outputformat4.setOption(str(self.txtOutputFormatFormatOptionKey_1_4.text()), str(self.txtOutputFormatFormatOptionValue_1_4.text()))
            if self.txtOutputFormatFormatOptionKey_2_4.text() != "" and self.txtOutputFormatFormatOptionValue_2_4.text() != "":
                outputformat4.setOption(str(self.txtOutputFormatFormatOptionKey_2_4.text()), str(self.txtOutputFormatFormatOptionValue_2_4.text()))
            if self.txtOutputFormatFormatOptionKey_3_4.text() != "" and self.txtOutputFormatFormatOptionValue_3_4.text() != "":
                outputformat4.setOption(str(self.txtOutputFormatFormatOptionKey_3_4.text()), str(self.txtOutputFormatFormatOptionValue_3_4.text()))
            if self.txtOutputFormatFormatOptionKey_4_4.text() != "" and self.txtOutputFormatFormatOptionValue_4_4.text() != "":
                outputformat4.setOption(str(self.txtOutputFormatFormatOptionKey_4_4.text()), str(self.txtOutputFormatFormatOptionValue_4_4.text()))
            ms_map.appendOutputFormat(outputformat4)
            
        # OWS metadata
        if self.txtMetadataOwsOwsTitle.text() != "":
            ms_map.setMetaData( "ows_title", self.txtMetadataOwsOwsTitle.text() )
             
        if self.txtMetadataOwsOwsOnlineResource.text() != "":
            ms_map.setMetaData( "ows_onlineresource", _toUtf8( u"%s?map=%s" % (self.txtMetadataOwsOwsOnlineResource.text(), self.txtMapFilePath.text()) ) )
        elif self.txtGeneralWebServerUrl.text() != "":
            ms_map.setMetaData( "ows_onlineresource", _toUtf8( u"%s?map=%s" % (self.txtGeneralWebServerUrl.text(), self.txtMapFilePath.text()) ) )
            
        srsList = []
        srsList.append( _toUtf8( self.canvas.mapRenderer().destinationCrs().authid() ) )
        if self.txtMetadataOwsOwsSrs.text() != "":
            ms_map.setMetaData( "ows_srs", self.txtMetadataOwsOwsSrs.text())
            
        if self.txtMetadataOwsOwsEnableRequest.text() != "":
            ms_map.setMetaData( "ows_enable_request", self.txtMetadataOwsOwsEnableRequest.text() )
#         if self.txtMetadataOwsGmlFeatureid.text() != "":
#             ms_map.setMetaData( "gml_featureid", self.txtMetadataOwsGmlFeatureid.text() )
#         if self.txtMetadataOwsGmlIncludeItems.text() != "":
#             ms_map.setMetaData( "gml_include_items", self.txtMetadataOwsGmlIncludeItems.text() )
        if self.txtMetadataWmsWebWmsEnableRequest.text() != "":
            ms_map.setMetaData( "wms_enable_request", self.txtMetadataWmsWebWmsEnableRequest.text() )
        if self.txtMetadataWfsWebWfsEnableRequest.text() != "":
            ms_map.setMetaData( "wfs_enable_request", self.txtMetadataWfsWebWfsEnableRequest.text() )
        if self.txtMetadataWcsWebWcsEnableRequest.text() != "":
            ms_map.setMetaData( "wcs_enable_request", self.txtMetadataWcsWebWcsEnableRequest.text() )
#             if self.txtMetadataWmsWebWmsEnableRequest.text() == "" and self.txtMetadataWfsWebWfsEnableRequest.text() == "" and self.txtMetadataWcsWebWcsEnableRequest.text() == "":
#                 ms_map.setMetaData( "ows_enable_request", "*" )
                
        if self.txtMetadataOwsWebOwsAllowedIpList.text() != "":
            ms_map.setMetaData( "ows_allowed_ip_list", self.txtMetadataOwsWebOwsAllowedIpList.text())
        if self.txtMetadataOwsWebOwsDeniedIpList.text() != "":
            ms_map.setMetaData( "ows_denied_ip_list", self.txtMetadataOwsWebOwsDeniedIpList.text())
        if self.txtMetadataOwsWebOwsSchemasLocation.text() != "":
            ms_map.setMetaData( "ows_schemas_location", self.txtMetadataOwsWebOwsSchemasLocation.text())
        if self.txtMetadataOwsWebOwsUpdatesequence.text() != "":
            ms_map.setMetaData( "ows_updatesequence", self.txtMetadataOwsWebOwsUpdatesequence.text())
        if self.txtMetadataOwsWebOwsHttpMaxAge.text() != "":
            ms_map.setMetaData( "ows_http_max_age", self.txtMetadataOwsWebOwsHttpMaxAge.text())
        if self.txtMetadataOwsWebOwsSldEnabled.text() != "":
            ms_map.setMetaData( "ows_sld_enabled", self.txtMetadataOwsWebOwsSldEnabled.text())
        
        # WMS metadata
        if self.txtMetadataWmsWebWmsTitle.text() != "":
            ms_map.setMetaData( "wms_title", self.txtMetadataWmsWebWmsTitle.text())
        if self.txtMetadataWmsWebWmsOnlineresource.text() != "":
            ms_map.setMetaData( "wms_onlineresource", _toUtf8( u"%s?map=%s" % (self.txtMetadataWmsWebWmsOnlineresource.text(), self.txtMapFilePath.text())))
        if self.txtMetadataWmsWebWmsSrs.text() != "":
            ms_map.setMetaData( "wms_srs", self.txtMetadataWmsWebWmsSrs.text())
        if self.txtMetadataWmsWebWmsAttrbutionOnlineresource.text() != "":
            ms_map.setMetaData( "wms_attribution_onlineresource", self.txtMetadataWmsWebWmsAttrbutionOnlineresource.text())
        if self.txtMetadataWmsWebWmsAttributionTitle.text() != "":
            ms_map.setMetaData( "wms_attribution_title", self.txtMetadataWmsWebWmsAttributionTitle.text())
        if self.txtMetadataWmsWebWmsBboxExtended.text() != "":
            ms_map.setMetaData( "wms_bbox_extended", self.txtMetadataWmsWebWmsBboxExtended.text())
        if self.txtMetadataWmsWebWmsFeatureInfoMimeType.text() != "":
            ms_map.setMetaData( "wms_feature_info_mime_type", self.txtMetadataWmsWebWmsFeatureInfoMimeType.text())
        if self.txtMetadataWmsWebWmsEncoding.text() != "":
            ms_map.setMetaData( "wms_encoding", self.txtMetadataWmsWebWmsEncoding.text())
        if self.txtMetadataWmsWebWmsGetcapabilitiesVersion.text() != "":
            ms_map.setMetaData( "wms_getcapabilities_version", self.txtMetadataWmsWebWmsGetcapabilitiesVersion.text())
        if self.txtMetadataWmsWebWmsFees.text() != "":
            ms_map.setMetaData( "wms_fees", self.txtMetadataWmsWebWmsTitle.text())
        if self.txtMetadataWmsWebWmsGetmapFormallist.text() != "":
            ms_map.setMetaData( "wms_getmap_formatlist", self.txtMetadataWmsWebWmsGetmapFormallist.text())
        if self.txtMetadataWmsWebWmsGetlegendgraphicFormailist.text() != "":
            ms_map.setMetaData( "wms_getlegendgraphic_formatlist", self.txtMetadataWmsWebWmsGetlegendgraphicFormailist.text())
        if self.txtMetadataWmsWebWmsKeywordlistVocabulary.text() != "":
            ms_map.setMetaData( "wms_keywordlist_vocabulary", self.txtMetadataWmsWebWmsKeywordlistVocabulary.text())
        if self.txtMetadataWmsWebWmsKeywordlist.text() != "":
            ms_map.setMetaData( "wms_keywordlist", self.txtMetadataWmsWebWmsKeywordlist.text())
        if self.txtMetadataWmsWebWmsLanguages.text() != "":
            ms_map.setMetaData( "wms_languages", self.txtMetadataWmsWebWmsLanguages.text())
        if self.txtMetadataWmsWebWmsLayerlimit.text() != "":
            ms_map.setMetaData( "wms_layerlimit", self.txtMetadataWmsWebWmsLayerlimit.text())
        if self.txtMetadataWmsWebWmsRootlayerKeywordlist.text() != "":
            ms_map.setMetaData( "wms_rootlayer_keywordlist", self.txtMetadataWmsWebWmsRootlayerKeywordlist.text())
        if self.txtMetadataWmsWebWmsRemoteSldMaxBytes.text() != "":
            ms_map.setMetaData( "wms_remote_sld_max_bytes", self.txtMetadataWmsWebWmsRemoteSldMaxBytes.text())
        if self.txtMetadataWmsWebWmsServiceOnlineResource.text() != "":
            ms_map.setMetaData( "wms_service_onlineresource", self.txtMetadataWmsWebWmsServiceOnlineResource.text())
        if self.txtMetadataWmsWebWmsResx.text() != "":
            ms_map.setMetaData( "wms_resx", self.txtMetadataWmsWebWmsResx.text())
        if self.txtMetadataWmsWebWmsResy.text() != "":
            ms_map.setMetaData( "wms_resy", self.txtMetadataWmsWebWmsResy.text())
        if self.txtMetadataWmsWebWmsRootlayerAbstract.text() != "":
            ms_map.setMetaData( "wms_rootlayer_abstract", self.txtMetadataWmsWebWmsRootlayerAbstract.text())
        if self.txtMetadataWmsWebWmsRootlayerTitle.text() != "":
            ms_map.setMetaData( "wms_rootlayer_title", self.txtMetadataWmsWebWmsRootlayerTitle.text())    
        if self.txtMetadataWmsWebWmsTimeformat.text() != "":
            ms_map.setMetaData( "wms_timeformat", self.txtMetadataWmsWebWmsTimeformat.text())
            
        # WFS metadata
        if self.txtMetadataWfsWebWfsTitle.text() != "":
            ms_map.setMetaData( "wfs_title", self.txtMetadataWfsWebWfsTitle.text())
        if self.txtMetadataWfsWebWfsOnlineresource.text() != "":
            ms_map.setMetaData( "wfs_onlineresource", _toUtf8( u"%s?map=%s" % (self.txtMetadataWfsWebWfsOnlineresource.text(), self.txtMapFilePath.text())))
        if self.txtMetadataWfsWebWfsSrs.text() != "":
            ms_map.setMetaData( "wfs_srs", self.txtMetadataWfsWebWfsSrs.text())
        if self.txtMetadataWfsWebWfsAbstract.text() != "":
            ms_map.setMetaData( "wfs_abstract", self.txtMetadataWfsWebWfsAbstract.text())
        if self.txtMetadataWfsWebWfsAccessconstraints.text() != "":
            ms_map.setMetaData( "wfs_accessconstraints", self.txtMetadataWfsWebWfsAccessconstraints.text())
        if self.txtMetadataWfsWebWfsEncoding.text() != "":
            ms_map.setMetaData( "wfs_encoding", self.txtMetadataWfsWebWfsEncoding.text())
        if self.txtMetadataWfsWebWfsFeatureCollection.text() != "":
            ms_map.setMetaData( "wfs_feature_collection", self.txtMetadataWfsWebWfsFeatureCollection.text())
        if self.txtMetadataWfsWebWfsFees.text() != "":
            ms_map.setMetaData( "wfs_fees", self.txtMetadataWfsWebWfsFees.text())
        if self.txtMetadataWfsWebWfsKeywordlist.text() != "":
            ms_map.setMetaData( "wfs_keywordlist", self.txtMetadataWfsWebWfsKeywordlist.text())
        if self.txtMetadataWfsWebWfsGetcapabilitiesVersion.text() != "":
            ms_map.setMetaData( "wfs_getcapabilities_version", self.txtMetadataWfsWebWfsGetcapabilitiesVersion.text())
        if self.txtMetadataWfsWebWfsNamespacePrefix.text() != "":
            ms_map.setMetaData( "wfs_namespace_prefix", self.txtMetadataWfsWebWfsNamespacePrefix.text())
        if self.txtMetadataWfsWebWfsMaxfeatures.text() != "":
            ms_map.setMetaData( "wfs_maxfeatures", self.txtMetadataWfsWebWfsMaxfeatures.text())
        if self.txtMetadataWfsWebWfsServiceOnlineresource.text() != "":
            ms_map.setMetaData( "wfs_service_onlineresource", self.txtMetadataWfsWebWfsServiceOnlineresource.text())
        if self.txtMetadataWfsWebWfsNamespaceUri.text() != "":
            ms_map.setMetaData( "wfs_namespace_uri", self.txtMetadataWfsWebWfsNamespaceUri.text())
            
        # WCS metadata
        if self.txtMetadataWcsWebWcsLabel.text() != "":
            ms_map.setMetaData( "wcs_label", self.txtMetadataWcsWebWcsLabel.text())
        if self.txtMetadataWcsWebWcsAbstract.text() != "":
            ms_map.setMetaData( "wcs_abstract", self.txtMetadataWcsWebWcsAbstract.text())
        if self.txtMetadataWcsWebWcsAccessconstraints.text() != "":
            ms_map.setMetaData( "wcs_accessconstraints", self.txtMetadataWcsWebWcsAccessconstraints.text())
        if self.txtMetadataWcsWebWcsDescription.text() != "":
            ms_map.setMetaData( "wcs_description", self.txtMetadataWcsWebWcsDescription.text())
        if self.txtMetadataWcsWebWcsKeywords.text() != "":
            ms_map.setMetaData( "wcs_keywords", self.txtMetadataWcsWebWcsKeywords.text())
        if self.txtMetadataWcsWebWcsFees.text() != "":
            ms_map.setMetaData( "wcs_fees", self.txtMetadataWcsWebWcsFees.text())
        if self.txtMetadataWcsWebWcsMetadatalinkFormat.text() != "":
            ms_map.setMetaData( "wcs_metadatalink_format", self.txtMetadataWcsWebWcsMetadatalinkFormat.text())
        if self.txtMetadataWcsWebWcsMetadatalinkHref.text() != "":
            ms_map.setMetaData( "wcs_metadatalink_href", self.txtMetadataWcsWebWcsMetadatalinkHref.text())
        if self.txtMetadataWcsWebWcsMetadatalinkType.text() != "":
            ms_map.setMetaData( "wcs_metadatalink_type", self.txtMetadataWcsWebWcsMetadatalinkType.text())
        if self.txtMetadataWcsWebWcsName.text() != "":
            ms_map.setMetaData( "wcs_name", self.txtMetadataWcsWebWcsName.text())
        if self.txtMetadataWcsWebWcsServiceOnlineresource.text() != "":
            ms_map.setMetaData( "wcs_service_onlineresource", self.txtMetadataWcsWebWcsServiceOnlineresource.text())
            

        layer_index = 0
#         
#         layers_array = [self.groupBoxLabelsLayer_1, self.groupBoxLabelsLayer_2, self.groupBoxLabelsLayer_3, self.groupBoxLabelsLayer_4]
#         layer_checkboxes_array = [self.checkBoxLayer_1, self.checkBoxLayer_2, self.checkBoxLayer_3, self.checkBoxLayer_4]
        
        for layer in self.legend.layers():
            
#             self.groupBoxLabelsLayer_1.SetChecked(True)
#             layers_array[layer_index].setEnabled(True)
            layer_index += 1
            # check if layer is a supported type... seems return None if type is not supported (e.g. csv)
            if ( self.getLayerType( layer ) == None):
                QMessageBox.warning(self, "RT MapServer Exporter", "Skipped not supported layer: %s" % layer.name())
                continue
            
            # create a layer object
            ms_layer = mapscript.layerObj( ms_map )
            ms_layer.name = _toUtf8( layer.name() )
            ms_layer.type = self.getLayerType( layer )
            ms_layer.status = self.onOffMap[ self.legend.isLayerVisible( layer ) ]
            ms_layer.dump = self.onOffMap[ True ]

            # layer extent
            extent = layer.extent()
            ms_layer.extent.minx = extent.xMinimum()
            ms_layer.extent.miny = extent.yMinimum()
            ms_layer.extent.maxx = extent.xMaximum()
            ms_layer.extent.maxy = extent.yMaximum()
            
            if self.txtMetadataOwsOwsEnableRequest.text() != "" or self.txtMetadataWfsWebWfsEnableRequest.text() != "":
                ms_layer.template = "dummy"

            ms_layer.setProjection( _toUtf8( layer.crs().toProj4() ) )

            if layer.hasScaleBasedVisibility():
                ms_layer.minscaledenom = layer.minimumScale()
                ms_layer.maxscaledenom = layer.maximumScale()

            # metadata from metadata config of each layer
            if self.checkBoxOws.isChecked():
                ms_layer.setMetaData( "ows_title", ms_layer.name )
            if self.checkBoxWms.isChecked():
                ms_layer.setMetaData( "wms_title", ms_layer.name )
            if self.checkBoxWfs.isChecked():
                ms_layer.setMetaData( "wfs_title", ms_layer.name )
                
            if self.txtMetadataOwsGmlFeatureid.text() != "":
                ms_layer.setMetaData( "gml_featureid", self.txtMetadataOwsGmlFeatureid.text() )
            if self.txtMetadataOwsGmlIncludeItems.text() != "":
                ms_layer.setMetaData( "gml_include_items", self.txtMetadataOwsGmlIncludeItems.text() )
                
            if layer.title() != "":
                ms_layer.setMetaData("wms_title", layer.title())
            if layer.abstract() != "":
                ms_layer.setMetaData("wms_abstract", layer.abstract())
            if layer.keywordList() != "":
                ms_layer.setMetaData("wms_keywordlist", layer.keywordList())
            if layer.dataUrl() != "":
                ms_layer.setMetaData("wms_dataurl_href", layer.dataUrl())
            if layer.dataUrlFormat() != "":
                ms_layer.setMetaData("wms_dataurl_format", layer.dataUrlFormat())
            if layer.attribution() != "":
                ms_layer.setMetaData("wms_attribution_title", layer.attribution())
            if layer.attributionUrl() != "":
                ms_layer.setMetaData("wms_attribution_onlineresource", layer.attributionUrl())
            if layer.metadataUrl() != "":
                ms_layer.setMetaData("wms_metadataurl_href", layer.metadataUrl())
            if layer.metadataUrlType() != "":
                ms_layer.setMetaData("wms_metadataurl_type", layer.metadataUrlType())
            if layer.metadataUrlFormat() != "":
                ms_layer.setMetaData("wms_metadataurl_format", layer.metadataUrlFormat())
            
            if layer.legendUrl() != "" or layer.legendUrlFormat() != "":
                layerLegenUrlStyle = ms_layer.name
                ms_layer.setMetaData("wms_style", layerLegenUrlStyle)
                if layer.legendUrl() != "":
                    ms_layer.setMetaData("wms_style_%s_legendurl_href" % layerLegenUrlStyle, layer.legendUrl())
                if layer.legendUrlFormat() != "":
                    ms_layer.setMetaData("wms_style_%s_legendurl_format" % layerLegenUrlStyle, layer.legendUrlFormat())
                if layer.legendUrl() != "":
                    ms_layer.setMetaData("wms_style_%s_legendurl_height" % layerLegenUrlStyle, "16")
                if layer.legendUrl() != "":
                    ms_layer.setMetaData("wms_style_%s_legendurl_width" % layerLegenUrlStyle, "16")
            

            # layer connection
            if layer.providerType() == 'postgres':
                ms_layer.setConnectionType( mapscript.MS_POSTGIS, "" )
                uri = QgsDataSourceURI( layer.source() )
                ms_layer.connection = _toUtf8( uri.connectionInfo() )
                data = u"%s FROM %s" % ( uri.geometryColumn(), uri.quotedTablename() )
                if uri.keyColumn() != "":
                    data += u" USING UNIQUE %s" % uri.keyColumn()
                data += u" USING srid=%s" % layer.crs().postgisSrid()
                if uri.sql() != "":
                  data += " FILTER (%s)" % uri.sql()
                ms_layer.data = _toUtf8( data )

            elif layer.providerType() == 'wms':
                ms_layer.setConnectionType( mapscript.MS_WMS, "" )

                uri = QUrl( "http://www.fake.eu/?"+layer.source() )
                ms_layer.connection = _toUtf8( uri.queryItemValue("url") )

                # loop thru wms sub layers
                wmsNames = []
                wmsStyles = []
                wmsLayerNames = layer.dataProvider().subLayers()
                wmsLayerStyles = layer.dataProvider().subLayerStyles()
                
                for index in range(len(wmsLayerNames)):
                    wmsNames.append( _toUtf8( wmsLayerNames[index] ) )
                    wmsStyles.append( _toUtf8( wmsLayerStyles[index] ) )

                # output SRSs
                srsList = []
                srsList.append( _toUtf8( layer.crs().authid() ) )

                # Create necessary wms metadata
                ms_layer.setMetaData( "ows_name", ','.join(wmsNames) )
                ms_layer.setMetaData( "wms_server_version", "1.1.1" )
                ms_layer.setMetaData( "ows_srs", ' '.join(srsList) )
                #ms_layer.setMetaData( "wms_format", layer.format() )
                ms_layer.setMetaData( "wms_format", ','.join(wmsStyles) )

            elif layer.providerType() == 'wfs':
                ms_layer.setConnectionType( mapscript.MS_WMS, "" )
                uri = QgsDataSourceURI( layer.source() )
                ms_layer.connection = _toUtf8( uri.uri() )

                # output SRSs
                srsList = []
                srsList.append( _toUtf8( layer.crs().authid() ) )

                # Create necessary wms metadata
                ms_layer.setMetaData( "ows_name", ms_layer.name )
                #ms_layer.setMetaData( "wfs_server_version", "1.1.1" )
                ms_layer.setMetaData( "ows_srs", ' '.join(srsList) )

            elif layer.providerType() == 'spatialite':
                ms_layer.setConnectionType( mapscript.MS_OGR, "" )
                uri = QgsDataSourceURI( layer.source() )
                ms_layer.connection = _toUtf8( uri.database() )
                ms_layer.data = _toUtf8( uri.table() )

            elif layer.providerType() == 'ogr':
                #ms_layer.setConnectionType( mapscript.MS_OGR, "" )
                ms_layer.data = _toUtf8( layer.source().split('|')[0] )

            else:
                ms_layer.data = _toUtf8( layer.source() )


            # set layer style
            if layer.type() == QgsMapLayer.RasterLayer:
                if hasattr(layer, 'renderer'):    # QGis >= 1.9
                    # layer.renderer().opacity() has range [0,1]
                    # ms_layer.opacity has range [0,100] => scale!
                    opacity = int( round(100 * layer.renderer().opacity()) )
                else:
                    opacity = int( 100 * layer.getTransparency() / 255.0 )
                ms_layer.opacity = opacity

            else:
                
                # use a SLD file set the layer style
                tempSldFile = QTemporaryFile("rt_mapserver_exporter-XXXXXX.sld")
                tempSldFile.open()
                tempSldPath = tempSldFile.fileName()
                tempSldFile.close()
                
                # export the QGIS layer style to SLD file
                errMsg, ok = layer.saveSldStyle( tempSldPath )
                if not ok:
                    QgsMessageLog.logMessage( errMsg, "RT MapServer Exporter" )
                else:
                    # set the mapserver layer style from the SLD file
                    #QFile.copy(tempSldPath, tempSldPath+".save")
                    #print "SLD saved file: ", tempSldPath+".save"
                    with open( unicode(tempSldPath), 'r' ) as fin:
                        sldContents = fin.read()
                    if mapscript.MS_SUCCESS != ms_layer.applySLD( sldContents, ms_layer.name ):
                        QgsMessageLog.logMessage( u"Something went wrong applying the SLD style to the layer '%s'" % ms_layer.name, "RT MapServer Exporter" )
                    QFile.remove( tempSldPath )
                
                    labelingEngine = QgsPalLabeling()
                    labelingEngine.loadEngineSettings()
    
                    if labelingEngine and labelingEngine.willUseLayer(layer):
                        ps = QgsPalLayerSettings.fromLayer(layer)
                        
                        if not ps.isExpression:
                            ms_layer.labelitem = unicode(ps.fieldName).encode('utf-8')
                        else:
                            pass
        
                        msLabel = mapscript.labelObj()
            
                        msLabel.type = mapscript.MS_TRUETYPE
                        msLabel.encoding = 'utf-8'
                        
                        # Position, rotation and scaling
                        msLabel.position = utils.serializeLabelPosition(ps)
                        msLabel.offsetx = int(ps.xOffset)
                        msLabel.offsety = int(ps.yOffset)
            
                        # Data defined rotation
                        # Please note that this is the only currently supported data defined property.
                        if QgsPalLayerSettings.Rotation in ps.dataDefinedProperties.keys():
                            dd = ps.dataDefinedProperty(QgsPalLayerSettings.Rotation)
                            rotField = unicode(dd.field()).encode('utf-8')
                            msLabel.setBinding(mapscript.MS_LABEL_BINDING_ANGLE, rotField)
                        else:
                            msLabel.angle = ps.angleOffset
            
                        if ps.scaleMin > 0:
                            ms_layer.labelminscaledenom = ps.scaleMin
                        if ps.scaleMax > 0:
                            ms_layer.labelmaxscaledenom = ps.scaleMax
            
                        fontDef, msLabel.size = utils.serializeFontDefinition(ps.textFont, ps.textNamedStyle)
            
                        # `emitFontDefinitions` gets set based on whether a fontset path is supplied through 
                        # the plugin UI. There is no point in emitting font definitions without a valid fontset,
                        # so in that case we fall back to using whatever default font MapServer (thus the
                        # underlying windowing system) provides.
                        # Please note that substituting the default font only works in MapServer 7.0.0 and above.
        #                 if emitFontDefinitions == True:
                        msLabel.font = fontDef
            
                        if ps.fontSizeInMapUnits:
                            utils.maybeSetLayerSizeUnitFromMap(QgsSymbolV2.MapUnit, ms_layer)
            
                        # Font size and color
                        msLabel.color = utils.serializeColor(ps.textColor)
            
                        if ps.fontLimitPixelSize:
                            msLabel.minsize = ps.fontMinPixelSize
                            msLabel.maxsize = ps.fontMaxPixelSize
            
                        # Other properties
                        wrap = unicode(ps.wrapChar).encode('utf-8')
                        if len(wrap) == 1:
                            msLabel.wrap = wrap[0]
                        elif len(wrap) > 1:
                            QgsMessageLog.logMessage(
                                u'Skipping invalid wrap character ("%s") for labels.' % wrap.decode('utf-8'),
                                'RT MapServer Exporter'
                            )
                        else:
                            # No wrap char set
                            pass
            
                        # Other properties
                        msLabel.partials = labelingEngine.isShowingPartialsLabels()
                        msLabel.force = ps.displayAll
                        msLabel.priority = ps.priority
                        msLabel.buffer = int(utils.sizeUnitToPx(
                            ps.bufferSize,
                            QgsSymbolV2.MapUnit if ps.bufferSizeInMapUnits else QgsSymbolV2.MM
                        ))
            
                        if ps.minFeatureSize > 0:
                            msLabel.minfeaturesize = utils.sizeUnitToPx(ps.minFeatureSize)
                            
                        ms_style = mapscript.styleObj()
                    
                        ms_style.color = mapscript.colorObj()
                        
                        ms_style.color.setRGB(ps.shapeFillColor.red(), ps.shapeFillColor.green(), ps.shapeFillColor.blue())
                        
                        ms_style.offsetx = float(ps.shapeOffset.x())
                        ms_style.offsety = float(ps.shapeOffset.y())
                        
                        if(ps.shapeSizeType == 0):
                            ms_style.setGeomTransform('labelpoly')
                        else:
#                         if(ps.shapeSizeType == SizeFixed):
                            ms_style.setGeomTransform('labelpnt')
            
                        # Label definitions gets appended to the very first class on a layer, or to a new class
                        # if no classes exist.
                        #
                        msLabel.insertStyle(ms_style)
                      
                        if ms_layer.numclasses > 0:
                            for c in range(0, ms_layer.numclasses):
                                ms_layer.getClass(c).addLabel(msLabel)
                        else:
                            mapscript.classObj(ms_layer).addLabel(msLabel)
                    
                

        # save the map file now!
        if mapscript.MS_SUCCESS != ms_map.save( _toUtf8( self.txtMapFilePath.text() )     ):
            return

        # Most of the following code does not use mapscript because it asserts
        # paths you supply exists, but this requirement is usually not meet on
        # the QGIS client used to generate the mafile.

        # get the mapfile content as string so we can manipulate on it
        mesg = "Reload Map file %s to manipulate it" % self.txtMapFilePath.text()
        QgsMessageLog.logMessage( mesg, "RT MapServer Exporter" )
        fin = open( _toUtf8(self.txtMapFilePath.text()), 'r' )
        parts = []
        line = fin.readline()
        while line != "":
            line = line.rstrip('\n')
            parts.append(line)
            line = fin.readline()
        fin.close()
        
        partsContentChanged = False

        # retrieve the list of used font aliases searching for FONT keywords
        fonts = []
        searchFontRx = re.compile("^\\s*FONT\\s+")
        for line in filter(searchFontRx.search, parts):
            # get the font alias, remove quotes around it
            fontName = re.sub(searchFontRx, "", line)[1:-1]
            # remove spaces within the font name
            fontAlias = fontName.replace(" ", "")

            # append the font alias to the font list
            if fontAlias not in fonts:
                fonts.append( fontAlias )

                # update the font alias in the mapfile
                # XXX: the following lines cannot be removed since the SLD file
                # could refer a font whose name contains spaces. When SLD specs
                # ate clear on how to handle fonts than we'll think whether
                # remove it or not.
                replaceFontRx = re.compile( u"^(\\s*FONT\\s+\")%s(\".*)$" % QRegExp.escape(fontName) )
                parts = [ replaceFontRx.sub(u"\g<1>%s\g<2>" % fontAlias, part) for part in parts ]
                partsContentChanged = True

        # create the file containing the list of font aliases used in the
        # mapfile
        if self.checkCreateFontFile.isChecked():
            fontPath = QFileInfo(_toUtf8(self.txtMapFilePath.text())).dir().filePath(u"fonts.txt")
            with open( unicode(fontPath), 'w' ) as fout:
                for fontAlias in fonts:
                    fout.write( unicode(fontAlias) )

        # add the FONTSET keyword with the associated path
        if self.txtMapFontsetPath.text() != "":
            # get the index of the first instance of MAP string in the list
            pos = parts.index( filter(lambda x: re.compile("^MAP(\r\n|\r|\n)*$").match(x), parts)[0] )
            if pos >= 0:
                parts.insert( pos+1, u'  FONTSET "%s"' % self.txtMapFontsetPath.text() )
                partsContentChanged = True
            else:
                QgsMessageLog.logMessage( u"'FONTSET' keyword not added to the mapfile: unable to locate the 'MAP' keyword...", "RT MapServer Exporter" )

        # if mapfile content changed, store the file again at the same path
        if partsContentChanged:
            with open( _toUtf8(self.txtMapFilePath.text()), 'w' ) as fout:
                for part in parts:
                    fout.write( part+"\n" )

        # XXX for debugging only: let's have a look at the map result! :)
        # XXX it works whether the file pointed by the fontset contains ALL the
        # aliases of fonts referred from the mapfile.
        #ms_map = mapscript.mapObj( unicode( self.txtMapFilePath.text() ) )
        #ms_map.draw().save( _toUtf8( self.txtMapFilePath.text() + ".png" )    , ms_map )

        QDialog.accept(self)


    def generateTemplate(self):
        tmpl = u""

        if self.getTemplateHeaderPath() == "":
            tmpl += u'''<!-- MapServer Template -->
<html>
  <head>
    <title>%s</title>
  </head>
  <body>
''' % self.txtGeneralMapName.text()

        for lid, orientation in self.templateTable.model().getObjectIter():
            layer = QgsMapLayerRegistry.instance().mapLayer( lid )
            if not layer:
                continue

            # define the template file content
            tmpl += '[resultset layer="%s"]\n' % layer.id()

            layerTitle = layer.title() if layer.title() != "" else layer.name()
            tmpl += u'<b>"%s"</b>\n' % layerTitle

            tmpl += '<table class="idtmplt_tableclass">\n'

            if orientation == Qt.Horizontal:
                tmpl += '  <tr class="idtmplt_trclass_1h">\n'
                for idx, fld in enumerate(layer.dataProvider().fields()):
                    fldDescr = fld.comment() if fld.comment() != "" else fld.name()
                    tmpl += u'    <td class="idtmplt_tdclass_1h">"%s"</td>\n' % fldDescr
                tmpl += '</tr>\n'

                tmpl += '[feature limit=20]\n'

                tmpl += '  <tr class="idtmplt_trclass_2h">\n'
                for idx, fld in enumerate(layer.dataProvider().fields()):
                    fldDescr = fld.comment() if fld.comment() != "" else fld.name()
                    tmpl += u'    <td class="idtmplt_tdclass_2h">[item name="%s"]</td>\n' % fld.name()
                tmpl += '  </tr>\n'

                tmpl += '[/feature]\n'

            else:
                for idx, fld in enumerate(layer.dataProvider().fields()):
                    tmpl += '  <tr class="idtmplt_trclass_v">\n'

                    fldDescr = fld.comment() if fld.comment() != "" else fld.name()
                    tmpl += u'    <td class="idtmplt_tdclass_1v">"%s"</td>\n' % fldDescr

                    tmpl += '[feature limit=20]\n'
                    tmpl += u'    <td class="idtmplt_tdclass_2v">[item name="%s"]</td>\n' % fld.name()
                    tmpl += '[/feature]\n'

                    tmpl += '  </tr>\n'

            tmpl += '</table>\n'

            tmpl += '[/resultset]\n'
            tmpl += '<hr>\n'


        if self.getTemplateFooterPath() == "":
            tmpl += '''  </body>
</html>'''

        return tmpl

    def getTemplatePath(self):
        if self.checkTmplFromFile.isChecked():
            return self.txtTemplatePath.text() # "[templatepath]"

        elif self.checkGenerateTmpl.isChecked():
            # generate the template for layers
            tmplContent = self.generateTemplate()
            # store the template alongside the mapfile
            tmplPath = self.txtMapFilePath.text() + ".html.tmpl"
            with open( unicode(tmplPath), 'w' ) as fout:
                fout.write( tmplContent )
            return tmplPath


class TemplateDelegate(QItemDelegate):
    """ delegate with some special item editors """

    def createEditor(self, parent, option, index):
        # special combobox for orientation
        if index.column() == 1:
            cbo = QComboBox(parent)
            cbo.setEditable(False)
            cbo.setFrame(False)
            for val, txt in enumerate(TemplateModel.ORIENTATIONS):
                cbo.addItem(txt, val)
            return cbo
        return QItemDelegate.createEditor(self, parent, option, index)

    def setEditorData(self, editor, index):
        """ load data from model to editor """
        m = index.model()
        if index.column() == 1:
            val = m.data(index, Qt.UserRole)[0]
            editor.setCurrentIndex( editor.findData(val) )
        else:
            # use default
            QItemDelegate.setEditorData(self, editor, index)

    def setModelData(self, editor, model, index):
        """ save data from editor back to model """
        if index.column() == 1:
            val = editor.itemData(editor.currentIndex())[0]
            model.setData(index, TemplateModel.ORIENTATIONS[val])
            model.setData(index, val, Qt.UserRole)
        else:
            # use default
            QItemDelegate.setModelData(self, editor, model, index)

class TemplateModel(QStandardItemModel):

    ORIENTATIONS = { Qt.Horizontal : u"Horizontal", Qt.Vertical : u"Vertical" }

    def __init__(self, parent=None):
        self.header = ["Layer name", "Orientation"]
        QStandardItemModel.__init__(self, 0, len(self.header), parent)

    def append(self, layer):
        rowdata = []

        item = QStandardItem( unicode(layer.name()) )
        item.setFlags( item.flags() & ~Qt.ItemIsEditable )
        rowdata.append( item )

        item = QStandardItem( TemplateModel.ORIENTATIONS[Qt.Horizontal] )
        item.setFlags( item.flags() | Qt.ItemIsEditable )
        rowdata.append( item )

        self.appendRow( rowdata )

        row = self.rowCount()-1
        self.setData(self.index(row, 0), layer.id(), Qt.UserRole)
        self.setData(self.index(row, 1), Qt.Horizontal, Qt.UserRole)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header[section]
        return None

    def getObject(self, row):
        lid = self.data(self.index(row, 0), Qt.UserRole)
        orientation = self.data(self.index(row, 1), Qt.UserRole)
        return (lid, orientation)

    def getObjectIter(self):
        for row in range(self.rowCount()):
            yield self.getObject(row)

