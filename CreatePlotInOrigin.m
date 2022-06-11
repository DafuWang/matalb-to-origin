function CreatePlotInOrigin()   
    originObj=actxserver('Origin.ApplicationSI'); %链接origin的COM接口
   
    invoke(originObj, 'Execute', 'doc -mc 1;');       
    invoke(originObj, 'IsModified', 'false');   

    invoke(originObj, 'Execute', 'syspath$=system.path.program$;');
    strPath='';
    strPath = invoke(originObj, 'LTStr', 'syspath$');
    invoke(originObj, 'Load', strcat(strPath, 'Samples\COM Server and Client\Matlab\CreatePlotInOrigin.OPJ'));
    
    mdata = [0.1:0.1:3; 10 * sin(0.1:0.1:3); 20 * cos(0.1:0.1:3)]; %创建数据
    mdata = mdata';  %将数据矩阵进行转置以适应worksheet列结构

    invoke(originObj, 'PutWorksheet', 'Data1', mdata);   %将数据传递至worksheet

    invoke(originObj, 'Execute', 'page.active = 1; layer - a; page.active = 2; layer - a;'); %图片绘制，创建2个图层
    invoke(originObj, 'CopyPage', 'Graph1'); % 图片复制到剪切板 
    
 end
