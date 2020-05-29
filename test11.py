def common_chart(chart_name,
                 chart_type,
                 chart_plotby,
                 sht=None,
                 chart_size=(0, 170, 900, 400),
                 data_rng=None,
                 chart_style_num=None,
                 no_gridline=True,
                 title_dict=None,
                 label_dict=None,
                 label_lines_color_dict=None,
                 legend_dict=None,
                 y_ticklabel_dict=None,
                 table_dict=None):
    '''
    c.xlColumnClustered --51

    c.xlColumns -- 2
    c.xlRows --1

    font_style_dict = {'font_style_dict':{'name':'Microsoft YaHei UI','size':24,'bold':True,'color':(255,255,255)}}
    pos_dict = {'pos':pos}
    series_style_dict = {'series_style_dict':
                            {'axis':2,'chart_type':c.,'smooth':False                                   }
                            }
                        }
    point_dict = {'s_num':1,'p_num':15}  + pos_dict + font_style + series_style_dict
    font_pos_dict = font_style + pos_dict
    label_lines_color_dict={'color_list':[255,255,255],'reverse':True}

    :param chart_name: 图对象 名称
    :param sht: 工作表
    :param data_rng: 数据区域
    :param chart_type: 图类型
    :param chart_size: 图大小 4长度的列表或元组
    :param chart_plotby: 横轴
    :param chart_style_num: 图款式
    :param no_gridline:无网格线
    :param title_dict:{'text':} + font_style_dict
    :param label_dict:point_dict
    :param legend_dict:font_pos_dict
    :param y_ticklabel_dict:font_style_dict
    :param table_dict:font_style_dict
    :return:
    '''
    if not sht:
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        sht = excel.ActiveSheet

    try:
        sht.ChartObjects(chart_name).Delete()
    except:
        pass
    finally:
        chart_obj = sht.Shapes.AddChart(chart_type,
                                        chart_size[0],
                                        chart_size[1],
                                        chart_size[2],
                                        chart_size[3])
        chart_obj.Name = chart_name
    chart = chart_obj.Chart

    if not data_rng:
        data_rng = sht.UsedRange

    chart.SetSourceData(Source=data_rng, PlotBy=chart_plotby)

    if chart_style_num:  # 整体格式
        chart.ChartStyle = chart_style_num

    if no_gridline:  # 无网格线
        chart.Axes(c.xlValue).HasMajorGridlines = False

    if title_dict:  # 标题
        chart.HasTitle = True
        title_obj = chart.ChartTitle
        title_obj.Text = title_dict['text']
        if 'font_style_dict' in title_dict:
            font_style(obj=title_obj, font_style_dict=title_dict['font_style_dict'])

    if label_dict:  # 标签
        series_list = []
        if isinstance(label_dict['s_num'], (list, tuple)):
            for _ in label_dict['s_num']:
                series_list.append(chart.FullSeriesCollection(_))
        elif isinstance(label_dict['s_num'], int):
            series_list.append(chart.FullSeriesCollection(label_dict['s_num']))
        elif label_dict['s_num'] == 'all':
            series_count = len(chart.FullSeriesCollection())
            for _ in range(1, series_count + 1):
                series_list.append(chart.FullSeriesCollection(_))

        for series in series_list:
            if 'series_style_dict' in label_dict:
                series_style_dict = label_dict['series_style_dict']
                if 'smooth' in series_style_dict:
                    series.Smooth = series_style_dict['smooth']
                if 'axis' in series_style_dict:
                    series.AxisGroup = series_style_dict['axis']
                if 'chart_type' in series_style_dict:
                    series.ChartType = series_style_dict['chart_type']

            if 'p_num' in label_dict:
                point_list = []
                if isinstance(label_dict['p_num'], (list, tuple)):
                    for _ in label_dict['p_num']:
                        point_list.append(series.Points(_))
                elif isinstance(label_dict['p_num'], int):
                    point_list.append(series.Points(label_dict['p_num']))
                elif label_dict['p_num'] == 'all':
                    points_count = len(series.Points())
                    for _ in range(1, points_count + 1):
                        point_list.append(series.Points(_))
                for _ in point_list:
                    if 'font_style_dict' in label_dict:
                        font_style(obj=_, font_style_dict=label_dict['font_style_dict'])
                    if 'pos_dict' in label_dict:
                        position(obj=_, pos=label_dict['pos'])

    if label_lines_color_dict:
        series_count = len(chart.FullSeriesCollection())
        position_list = list(range(1, series_count + 1))
        series_lines_color_list = label_lines_color_dict['color_list']
        if label_lines_color_dict['reverse']:
            series_lines_color_list = series_lines_color_list[::-1]
            position_list = position_list[::-1]
        for n, i in enumerate(position_list):
            series = chart.FullSeriesCollection(i)
            series.Format.Line.ForeColor.RGB = rgbToInt(series_lines_color_list[n])

    if legend_dict:
        chart.HasLegend = True
        legend_obj = chart.Legend
        if 'font_style_dict' in legend_dict:
            font_style(obj=legend_obj, font_style_dict=legend_dict['font_style_dict'])
        if 'pos' in legend_dict:
            position(obj=legend_obj, pos=legend_dict['pos'])

    if y_ticklabel_dict:
        y_ticklabel = chart.Axes(c.xlValue).TickLabels
        if 'font_style_dict' in y_ticklabel_dict:
            font_style(obj=y_ticklabel, font_style_dict=y_ticklabel_dict['font_style_dict'])

    if table_dict:
        chart.HasDataTable = True
        table_obj = chart.DataTable
        if 'font_style_dict' in y_ticklabel_dict:
            font_style(obj=table_obj, font_style_dict=y_ticklabel_dict['font_style_dict'])

    chart_obj.Select()
    return chart_obj