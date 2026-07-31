                df_c["Label"] = df_c["Downtime name"]
                colors_map = None
            fig = px.bar(
                df_c, x="confessed [h]", y="Label", orientation="h",
                color="Label" if colors_map else None,
                color_discrete_map=colors_map if colors_map else None,
                color_discrete_sequence=[APTIV_ACCENT] if not colors_map else None,
            )
            fig.update_traces(
                marker=dict(cornerradius=4),
                texttemplate="%{x:.2f} h", textposition="outside",
                textfont=dict(size=10, color=MUTED),
                hovertemplate="<b>%{y}</b><br>%{x:.2f} h<extra></extra>",
            )
            _style_fig(fig, height=310, legend=False)
            fig.update_layout(xaxis_title="Heures d'arrêt", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Aucun arrêt enregistré pour la sélection actuelle.")
    else:
        st.info("Pas de données disponibles.")

section("Contribution par machine", "Heures d'arrêt empilées par catégorie")
if not raw_log_f.empty and not dt_legend.empty and "CodeDT_num" in dt_legend.columns:
    raw_with_cat_m = raw_log_f.merge(
        dt_legend[["CodeDT_num", "Categorie", "ColorHEX"]],
        left_on="Downtime reason", right_on="CodeDT_num", how="left",
    )
    raw_with_cat_m["Categorie"] = raw_with_cat_m["Categorie"].fillna("Non catégorisé")
    only_dt = raw_with_cat_m[raw_with_cat_m["Downtime reason"] != 0]

    df_mc = only_dt.groupby(["Machine", "Categorie"])["confessed [h]"].sum().reset_index()
    if not df_mc.empty and df_mc["confessed [h]"].sum() > 0:
        cat_color_map_m = (
            raw_with_cat_m.dropna(subset=["Categorie"])
            .drop_duplicates(subset=["Categorie"])
            .set_index("Categorie")["ColorHEX"].to_dict()
        )
        machine_order = (
            df_mc.groupby("Machine")["confessed [h]"].sum()
            .sort_values(ascending=False).index.tolist()
        )
        fig_m = px.bar(
            df_mc, x="Machine", y="confessed [h]", color="Categorie",
            category_orders={"Machine": machine_order},
            color_discrete_map=cat_color_map_m,
            custom_data=["Categorie"],
        )
        fig_m.update_traces(
            hovertemplate="<b>%{x}</b><br>%{customdata[0]} : %{y:.2f} h<extra></extra>",
        )
        _style_fig(fig_m, height=320)
        fig_m.update_layout(barmode="stack", xaxis_title="", yaxis_title="Heures d'arrêt",
                            legend=dict(orientation="h", y=-0.25, font=dict(size=10)))
        event_m = st.plotly_chart(
            fig_m, use_container_width=True,
            on_select="rerun", selection_mode="points", key="machine_cat_chart",
        )

        points = (event_m or {}).get("selection", {}).get("points", [])
        if points:
            p = points[0]
            machine_sel = p.get("x")
            cat_sel = None
            if "customdata" in p and p["customdata"]:
                cat_sel = p["customdata"][0]
            detail = only_dt[
                (only_dt["Machine"] == machine_sel) & (only_dt["Categorie"] == cat_sel)
            ][["Machine", "Downtime name", "Commentaire", "confessed [h]"]]
            st.markdown(f"**Détail — Machine `{machine_sel}` / catégorie `{cat_sel}`**")
            if not detail.empty:
                st.dataframe(detail.sort_values("confessed [h]", ascending=False),
                             use_container_width=True, hide_index=True)
            else:
                st.caption("Aucun détail trouvé pour cette sélection.")
        else:
            st.caption("Sélectionnez une barre pour afficher le détail des arrêts correspondants.")
    else:
        st.info("Aucun arrêt enregistré pour la sélection actuelle.")
else:
    st.info("Légende DT_LEGEND indisponible pour catégoriser les arrêts.")

with st.expander("Consulter les données détaillées filtrées"):
    if not raw_log_f.empty:
        st.markdown("**RAW_LOG**")
        st.dataframe(raw_log_f, use_container_width=True)
    if not summary_f.empty:
        st.markdown("**SUMMARY**")
        st.dataframe(summary_f, use_container_width=True)
    if not dt_legend.empty:
        st.markdown("**DT_LEGEND**")
        st.dataframe(dt_legend, use_container_width=True)

st.markdown(
    f"""
    <div class="app-footer">
      <div>APTIV · Operations Excellence — Dashboard OEE interne</div>
      <div>{len(files)} fichier(s) source · généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</div>
    </div>
    """,
    unsafe_allow_html=True,
)
