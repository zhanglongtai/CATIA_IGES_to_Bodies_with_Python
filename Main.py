# -*- coding:utf-8 -*-
import win32com.client
import os

catApp = win32com.client.Dispatch('CATIA.Application')
catApp.Visible = True

all_file = os.listdir('E:\\barobot_CATIA\\')

for i in all_file:
    
    suffix = i.split('.')
    suffix = suffix[-1].lower()
    
    if suffix == 'igs':
        
        try:
            
            Model = catApp.Documents.Open('E:\\barobot_CATIA\\' + i)
    
            Model_Part = catApp.ActiveDocument.Part
    
            Model_Bodies = Model_Part.Bodies
    
            Model_HybridBodies = Model_Part.HybridBodies
    
            Model_HSF = Model_Part.HybridShapeFactory
            Model_SF = Model_Part.ShapeFactory
    
            Model_Geometrical_Set = Model_HybridBodies.Item(1)
    
            Model_GS_Element = Model_Geometrical_Set.HybridShapes
    
            Element_Number = Model_GS_Element.Count
    
            if Element_Number <= 2:
                if Element_Number > 1:
                    l = [ Model_GS_Element.Item(1),  Model_GS_Element.Item(2)]
                    JoinFaces = Model_HSF.AddNewJoin(l[0], l[1])
                else:
                    JoinFaces = Model_GS_Element.Item(1)
            else:
                l = []
                for n in range(Element_Number):
                    l.append(Model_GS_Element.Item(n + 1))
                JoinFaces = Model_HSF.AddNewJoin(l[0], l[1])
                for n in range(Element_Number - 2):
                    JoinFaces.AddElement(l[n + 2])
    
            Model_Geometrical_Set.AppendHybridShape(JoinFaces)
    
            Model_Part.Update()
    
            Model_Part.InWorkObject = Model_Bodies.Item(1)
    
            Model_SF.AddNewCloseSurface(JoinFaces)
            Model_HSF.GSMVisibility(JoinFaces, 0)
            Model_Part.Update()
            
            Name = i.split('.')[:-1]
            Name = Name[0].split()
            FileName = ''
            for word in Name:
                if word == Name[-1]:
                    FileName += word.capitalize() + '.CATPart'
                else:
                    FileName += word.capitalize() + '_'
            
            catApp.ActiveDocument.SaveAs('E:\\barobot_CATIA\\' + FileName)
            Model.Close()
        
        except:
            print i + ' has encountered problems.'
