# import required modules 
from math import e
import requests, json 
import geopandas as gpd
import pandas as pd
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import folium
import plotly.express as px
import seaborn as sns
import pandapower as pp
import pandapower.shortcircuit as sc
import pandapower.plotting.plotly as plty
from pandapower.plotting.plotly import simple_plotly, pf_res_plotly, vlevel_plotly
import geopandas as gpd
import contextily as ctx
import streamlit as st
from streamlit.elements.media import YOUTUBE_RE
import base64
from io import BytesIO


# Streamlit App Title
st.title("Geospatial Power System Analysis")
st.markdown("""
This application demosntrates a web-based geospatial power flow analysis 
.The application makes use of actual data from a distribution feeder in Nairobi,Kenya.
It makes use of the followign resurces.
* **PandaPower Library** : Pandapower combines the data analysis library pandas and
the power flow solver PYPOWER to create a power systems analysis package for
automation of analysis and optimization in power systems.
Aside from that, it allows for geospatial plotting of te network, as long as 
network geodata is provided.
""")
st.sidebar.info("Select and Enter Calculation Parameters")
#####################################################################
@st.cache(allow_output_mutation=True)
def create_net():
    bus=pd.read_excel('line_data.xlsx',sheet_name='bus',index_col=0)
    trafos=pd.read_excel('line_data.xlsx',sheet_name='trafos',index_col=0)
    lines=pd.read_excel('line_data.xlsx',sheet_name='lines',index_col=0)
    loads=pd.read_excel('line_data.xlsx',sheet_name='load',index_col=0)
    net=pp.create_empty_network()
    bus.head()

    #####################################################################
    for idx in bus.index:
        pp.create_bus(net,name=bus.at[idx,'bus_name'],
                    vn_kv=bus.at[idx,'v_nom'],
                    in_service='TRUE',geodata=(bus.at[idx,'x'],bus.at[idx,'y']))
    #####################################################################
    for idx in loads.index:
        pp.create_load(net,name=loads.at[idx,"name"],bus=loads.at[idx,"bus"],in_service=loads.at[idx,"in_service"],
                    p_mw=loads.at[idx,"p_set"],q_mvar=loads.at[idx,"q_set"])

    #####################################################################

    for idx in lines.index:
        pp.create_line_from_parameters(net,name=lines.at[idx,'name'],from_bus = lines.at[idx,'from_bus'], 
                                    to_bus = lines.at[idx,'to_bus'],
                                    length_km=lines.at[idx,'length'],
                                    r_ohm_per_km = lines.at[idx,'r'],
                                    x_ohm_per_km = lines.at[idx,'x'],
                                    c_nf_per_km = lines.at[idx,'b'],
                                    max_i_ka = lines.at[idx,'max_i'],
                                    parallel=lines.at[idx,'parallel'],coords=lines.at[idx,'geometry'])

    #####################################################################
    pl_mwtrafos=trafos.dropna()
    for idx in trafos.index:
        pp.create_transformer_from_parameters(net,name=trafos.at[idx,'name'], hv_bus= trafos.at[idx,'from_bus'], lv_bus=trafos.at[idx,'to_bus'],
                                            sn_mva=trafos.at[idx,'s_nom'],vn_hv_kv=trafos.at[idx,'vn_hv_kv'],                                          
                                                vn_lv_kv=trafos.at[idx,'vn_lv_kv'], vkr_percent=trafos.at[idx,'vkr_percent'],
                                            vk_percent=trafos.at[idx,'vk_percent'],
                                            pfe_kw=trafos.at[idx,'pfe'], 
                                                i0_percent=trafos.at[idx,'i0_percent'], shift_degree=0,
                                                tap_phase_shifter=False, in_service=True,
                                                index=idx, max_loading_percent=80, parallel=1, df=1.0)

    ##########################################################################
    pp.create_ext_grid(net, bus=101, vm_pu=1.0, va_degree=0.0, name="ext_grid")
    return net
net=create_net()

#####################################################################

Calculations=['Power Flow Analysis','Short Circuit Analysis']
CalcSelect=st.sidebar.radio("Select Calculation",Calculations)

#####################################################################

Algorithms={"nr":"Newton-Raphson",'iwamoto_nr':'Iwamoto_nr','gs':'Gauss_Siedel',
'bfsw':'Backward/Forward Sweep','fdbx':'Fast-Decoupled'}

# Initialization={"Auto":'auto','FlatStart':'flat','DC':'dc',
# 'results':'results'}

Initialization={'auto':"Auto",'flat':'FlatStart','dc':'DC',
'results':'results'}



ReactiveLim=[True,False]

#ShortCircuit Params
FaultType=["1ph",'2ph','3ph']
Case=['max','min']

if CalcSelect == "Power Flow Analysis":
    Algoselect=st.sidebar.selectbox("Select Algorithm",list(Algorithms.keys()),format_func=lambda x: Algorithms.get(x))#options=list(Algorithms.keys()))
    InitSelect=st.sidebar.selectbox("Select Initialization Method",list(Initialization.keys()),format_func=lambda x: Initialization.get(x))
    QlimSelect=st.sidebar.selectbox("Enforce Qlimits?",ReactiveLim)
    MaxIter=st.sidebar.number_input("Enter Maximum Iterations",min_value=0,max_value=100,step=1)
    RunLoadFlow=st.sidebar.button("Calculate")
    if RunLoadFlow:
        pp.runpp(net,algorithm= Algoselect,init=InitSelect,enforce_q_lims=QlimSelect)
 ####################################################################
 # Download Results
        @st.cache
        def create_res_excel():
            bus_res=pd.merge(net.bus,net.res_bus,left_index=True, right_index=True)
            bus_res=pd.merge(bus_res,net.bus_geodata,left_index=True, right_index=True)
            lines_res=pd.merge(net.line,net.res_line,left_index=True, right_index=True)
            trafo_res=pd.merge(net.trafo,net.res_trafo,left_index=True, right_index=True)
            load_res=pd.merge(net.load,net.res_load,left_index=True, right_index=True)
            ext_grid_res=pd.merge(net.ext_grid,net.res_ext_grid,left_index=True, right_index=True)
            # Create a Pandas Excel writer using XlsxWriter as the engine.
            output=BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            # Write each dataframe to a different worksheet.
            bus_res.to_excel(writer, sheet_name='Bus_Results')
            lines_res.to_excel(writer, sheet_name='Line_Results')
            trafo_res.to_excel(writer, sheet_name='Trafo_Results')
            load_res.to_excel(writer, sheet_name='Load_Results')
            ext_grid_res.to_excel(writer, sheet_name='Ext_Grid Results')
            # Close the Pandas Excel writer and output the Excel file.
            writer.save()
            processed_data = output.getvalue()
            return processed_data 
        processed_data =create_res_excel()
        st.download_button(
        label="Download LoadFlow Results File",
        data=processed_data ,
        file_name= 'LoadFlowResults.xlsx'
        
 )

        col1,col2,col3,col4=st.columns(4)
        with col1:
            st.metric("Vmax(pu)",round(net.res_bus['vm_pu'].max(),2))
        with col2:
            st.metric("Vmin(pu)",round(net.res_bus['vm_pu'].min(),2))
        with col3:
            st.metric("Ploss(%)",round(net.res_line['pl_mw'].sum()/net.res_ext_grid['p_mw'].sum(),2))
        with col4:
            st.metric("Total Import/Export(MW)",round(net.res_ext_grid['p_mw'].sum(),2))

        figz=pf_res_plotly(net, on_map=True, projection='epsg:4326', map_style='satellite')
        st.plotly_chart(figz,use_container_width=True)
        colus,colus1=st.columns(2)
        with colus:
            fig = px.histogram(net.res_bus, x='vm_pu',title="Bus Voltage Distr.",color_discrete_sequence=['red'])
            st.plotly_chart(fig,use_container_width=True)
        with colus1:
            fig = px.box(net.res_trafo, y='loading_percent',title="Transformer Loading Distr.",color_discrete_sequence=['goldenrod'])
            st.plotly_chart(fig,use_container_width=True)


    else:
        st.sidebar.write("Press To Calculate Load Flow")

#####################################################################


else:
    SelFaultType=st.sidebar.selectbox("Select Fault Type",FaultType)#options=list(Algorithms.keys()))
    SelectCaseType=st.sidebar.selectbox("Select Fault Case",Case)
    
    # if Case='min':
    ExtGridMVAInput=st.sidebar.number_input("Ext.Grid Short Cct MVA")
    ExtGridrxmin=st.sidebar.number_input("Ext.Grid rx_min",min_value=0.0,max_value=1.0,value=0.0)
    Temp_Rise=st.sidebar.number_input("Max Temp Rise",min_value=0.0,value=0.0)
    net.ext_grid["s_sc_min_mva"] = ExtGridMVAInput
    net.ext_grid["s_sc_max_mva"] = ExtGridMVAInput
    net.ext_grid["rx_min"] = ExtGridrxmin
    net.ext_grid["rx_max"] = ExtGridrxmin
    net.line["endtemp_degree"] = Temp_Rise
    RunSCCalc=st.sidebar.button("Calculate Short Circuits")
    if RunSCCalc:
        sc.calc_sc(net, case=SelectCaseType,fault=SelFaultType)

        col1,col2=st.columns(2)
        with col1:
            st.metric("Max Short Circuit",round(net.res_bus_sc['ikss_ka'].max(),2))
        with col2:
            st.metric("Min Short Circuit Current",round(net.res_bus_sc['ikss_ka'].min(),2))

        colus,colus1=st.columns(2)
        with colus:
            fig = px.histogram(net.res_bus_sc, x='ikss_ka',title="Short Circuit Current Distribution",color_discrete_sequence=['red'])
            st.plotly_chart(fig,use_container_width=True)
        with colus1:
            fig = px.box(net.res_bus_sc, y='skss_mw',title="Short Circuit MVA Distribution",color_discrete_sequence=['goldenrod'])
            st.plotly_chart(fig,use_container_width=True)
      
        bus_res_sc=pd.merge(net.bus,net.res_bus_sc,left_index=True, right_index=True)
        bus_res_sc=pd.merge(bus_res_sc,net.bus_geodata,left_index=True, right_index=True)
        for i in bus_res_sc['vn_kv']:
            if i ==bus_res_sc['vn_kv'].unique()[0]:
                bus_res_sc['ShortCctRating']=6
            elif i ==bus_res_sc['vn_kv'].unique()[1] :
                bus_res_sc['ShortCctRating']=31.5
            else:
                bus_res_sc['ShortCctRating']=0
                
        bus_res_sc['PercShortCctRating']=bus_res_sc["ikss_ka"]/bus_res_sc['ShortCctRating']
       
        bus_res_sc=bus_res_sc.sort_values(by='PercShortCctRating',ascending=False)

        figz2 = px.bar(bus_res_sc[:10], x='name',y='PercShortCctRating',
        title="Ik Percentage of Switchgear Rating Per Bus (Top 10)",
        labels={"name":"BusName","PercShortCctRating":"% Switchgear Rating"})
        figz2.update_geos(fitbounds="locations", visible=False)
        st.plotly_chart(figz2,use_container_width=True)
        st.balloons()

    else:
        st.sidebar.write("Press Button to Calculate Short Circuits")
    


#####################################################################
