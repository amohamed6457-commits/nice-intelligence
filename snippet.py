    if "Clinical endpoints" in shown and agency == "HAS":
        ep = load_has_endpoints(version)
        if ep.empty:
            st.caption("has_endpoints_full.json not found next to app3.py.")
        else:
            ids = frame["appraisal_id"] if "appraisal_id" in frame else []
            ov = hep.overview(ep, ids)
            if not ov:
                st.caption("No extracted endpoints for this selection.")
            else:
                st.markdown("**Clinical endpoints reported by the committee**")
                a, b = st.columns(2)
                a.metric("Endpoints", f"{ov['endpoints']:,}")
                b.metric("Appraisals covered", f"{ov['appraisals']:,}")
                c1, c2 = st.columns(2)
                with c1:
                    st.dataframe(ov["categories"], hide_index=True,
                                 width="stretch")
                with c2:
                    st.dataframe(ov["commonest"], hide_index=True,
                                 width="stretch")
                st.caption(hep.CAVEAT)
