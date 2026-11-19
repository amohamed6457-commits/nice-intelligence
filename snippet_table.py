    if "Clinical endpoints" in shown and agency == "HAS":
        st.markdown("**Clinical endpoints reported by the committee**")
        hep.render_endpoint_table(st, load_has_endpoints(version),
                                  load_has_endpoint_links(version), frame)
