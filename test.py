            with tab3:
                st.markdown('<div class="stl c">🔬 Analyse Detaillee</div>', unsafe_allow_html=True)

                # Définition de la marge pour éviter les erreurs de frappe
                mg = dict(t=40, b=10, l=10, r=10)

                st.markdown('<div class="stl p" style="margin-top:8px">📦 Analyse Backlog Preparation</div>', unsafe_allow_html=True)
                bl_prep_data = dfp[dfp["Statut OT"]=="CRÉÉ"].copy()
                bl_prep_car = bl_prep_data[bl_prep_data["Backlog preparation"]=="CARACTERISE"]
                bl_prep_ncar = bl_prep_data[bl_prep_data["Backlog preparation"]=="NON CARACTERISE"]
                bp_prep_total = len(bl_prep_data)
                if bp_prep_total > 0:
                    c1, c2 = st.columns(2)
                    with c1:
                        bp_prep_pie = pd.DataFrame({
                            "Statut": ["Caracterise (%d)" % len(bl_prep_car), "Non Caracterise (%d)" % len(bl_prep_ncar)],
                            "Nombre": [len(bl_prep_car), len(bl_prep_ncar)]
                        })
                        fig_pp = anl_pie_chart(bp_prep_pie, "Statut", "Nombre", "Backlog Preparation", ["#38a169","#e53e3e"])
                        if fig_pp: st.plotly_chart(fig_pp, use_container_width=True)
                    with c2:
                        prep_by_poste = bl_prep_data.groupby("Poste travail princ.").size().sort_values(ascending=False).head(10)
                        prep_car_by_poste = bl_prep_car.groupby("Poste travail princ.").size()
                        prep_df = pd.DataFrame({
                            "Poste": prep_by_poste.index,
                            "Total": prep_by_poste.values,
                            "Caracterise": [prep_car_by_poste.get(p,0) for p in prep_by_poste.index]
                        })
                        fig_pb = px.bar(prep_df, x="Total", y="Poste", orientation="h",
                                        title="Top 10 Postes - Backlog Preparation",
                                        color="Caracterise", color_discrete_sequence=["#38a169","#e53e3e"])
                        fig_pb.update_layout(height=450, autosize=True, margin=mg, title_font_size=11)
                        st.plotly_chart(fig_pb, use_container_width=True)
                else:
                    st.markdown('<div class="es">Aucun OT en statut CRÉÉ pour cette periode</div>', unsafe_allow_html=True)

                st.markdown('<div class="stl p" style="margin-top:8px">⏳ Analyse Age des OT par Statut</div>', unsafe_allow_html=True)
                for statut_label, statut_filter, age_col, kpi_list_age in [
                    ("Preparation (CRÉÉ)", dfp["Statut OT"]=="CRÉÉ", "ap",
                     ["OT préparation <1 mois","OT préparation 1mois< <3mois","OT préparation >3 mois"]),
                    ("Planification (LANC sans SOPL)", (dfp["Statut OT"]=="LANC")&(dfp["Contient SOPL"]==0), "alp",
                     ["OT planification <1 mois","OT planification 1mois< <3mois","OT planification >3 mois"]),
                    ("Execution (LANC avec SOPL)", (dfp["Statut OT"]=="LANC")&(dfp["Contient SOPL"]==1), "aex",
                     ["OT exécution <1 mois","OT exécution 1mois< <3mois","OT exécution >3 mois"])
                ]:
                    sub = dfp[statut_filter]
                    if len(sub) > 0:
                        age_dist = sub[age_col].value_counts()
                        age_data = pd.DataFrame({"Categorie": age_dist.index, "Nombre": age_dist.values})
                        fig_age = px.pie(age_data, names="Categorie", values="Nombre",
                                         title=f"Repartition Age - {statut_label}",
                                         color_discrete_sequence=["#38a169","#ecc94b","#e53e3e"])
                        fig_age.update_traces(textposition='inside', textinfo='percent+label+value', textfont_size=9)
                        fig_age.update_layout(height=450, autosize=True, margin=mg, title_font_size=11,
                                              legend=dict(font_size=8, orientation="h", yanchor="bottom", y=-0.15))
                        st.plotly_chart(fig_age, use_container_width=True)

                if all_ano:
                    st.markdown('<div class="stl a" style="margin-top:8px">🔍 Detail des Anomalies</div>', unsafe_allow_html=True)
                    ano_df = pd.DataFrame(all_ano)
                    if desig_col and desig_col in dfp.columns:
                        detail_rows = []
                        for _, ar in ano_df.iterrows():
                            poste = ar["Poste"]; kpi = ar["KPI"]
                            dp = dfp[dfp["Poste travail princ."]==poste]
                            sub_fn = sub_p.get(kpi) or sub_q.get(kpi)
                            if sub_fn:
                                anomalies = sub_fn(dp)
                                for _, arow in anomalies.head(5).iterrows():
                                    detail_rows.append({
                                        "Poste": poste, "KPI": kpi,
                                        "OT": arow.get("Ordre",""),
                                        "Designation": str(arow.get(desig_col,""))[:80],
                                        "Statut": arow.get("Statut OT","")
                                    })
                        if detail_rows:
                            ddf = pd.DataFrame(detail_rows)
                            st.markdown(anl_html_table(ddf), unsafe_allow_html=True)
