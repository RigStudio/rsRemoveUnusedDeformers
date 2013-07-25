##
# rsRemoveUnusedDeformers
# @author Juan Lara
# @date 2013-07-25
# @file rsRemoveUnusedDeformers.py

import win32com.client
from win32com.client import constants

Application = win32com.client.Dispatch('XSI.Application').Application


def XSILoadPlugin(in_reg):
    in_reg.Author = 'jlara'
    in_reg.Name = 'rsRemoveUnusedDeformers'
    in_reg.Email = 'info@rigstudio.com'
    in_reg.URL = 'www.rigstudio.com'
    in_reg.Major = 1
    in_reg.Minor = 0

    in_reg.RegisterCommand('rsRemoveUnusedDeformers', 'rsRemoveUnusedDeformers')
    in_reg.RegisterMenu(constants.siMenuTbAnimateDeformEnvelopeID, 'rsRemoveUnusedDeformers_Menu', False, False)

    return True


def XSIUnloadPlugin(in_reg):
    strPluginName = in_reg.Name
    Application.LogMessage(str(strPluginName) + str(' has been unloaded.'), constants.siVerbose)
    return True


def rsRemoveUnusedDeformers_Menu_Init(ctxt):
    oMenu = ctxt.Source
    oMenu.AddCommandItem('rsRemoveUnusedDeformers', 'rsRemoveUnusedDeformers')
    return True


def rsRemoveUnusedDeformers_Init(in_ctxt):
    o_cmd = in_ctxt.Source
    o_cmd.Description = ''
    o_cmd.ReturnValue = True

    o_args = o_cmd.Arguments
    o_args.AddWithHandler('in_c_mesh', 'Collection')

    return True


def rsRemoveUnusedDeformers_Execute(in_c_mesh):
    
    if in_c_mesh.Count == 0:
        Application.LogMessage('Nothing Selected', 2)
        return False

    c_remove = win32com.client.Dispatch('XSI.Collection')

    for o_mesh in in_c_mesh:
        o_mesh = get3DObject(o_mesh)

        for o_envelope in o_mesh.Envelopes:
            c_deformers = o_envelope.Deformers
            l_weights = o_envelope.Weights.Array
            c_remove.RemoveAll()
    
            for i_deformer in range(len(l_weights)):
                o_deformer = c_deformers(i_deformer)
                if sum(l_weights[i_deformer]) == 0:
                    if not o_deformer.Type in ['eff', 'root']:
                        c_remove.Add(o_deformer)
    
            if c_remove.Count != 0:
                Application.RemoveFlexEnvDeformer('%s;%s' % (o_mesh.FullName, c_remove.GetAsText()), False)
    
    return True


def get3DObject(in_o_obj):
    d_type = {'polySubComponent': 1, 'edgeSubComponent': 1, 'pntSubComponent': 1, 'sampleSubComponent': 1,
              'poly': 2, 'edge': 2, 'pnt': 2}
    
    i_action = d_type.get(in_o_obj.Type, 3)
    if i_action == 1:
        o_out = in_o_obj.SubComponent.Parent3DObject
    elif i_action == 2:
        o_out = in_o_obj.Parent3DObject
    else:
        o_out = in_o_obj

    return o_out
