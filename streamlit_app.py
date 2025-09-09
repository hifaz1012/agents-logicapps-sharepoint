import streamlit as st
import requests

st.title("SharePoint Agent Logic Apps UI")

user_query = st.text_area("Enter your query for the agent:")

if st.button("Run Agent"):
    if user_query.strip():
        try:
            response = requests.post(
                "http://localhost:8000/run-agent",
                json={"query": user_query}
            )
            if response.status_code == 200:
                results = response.json().get("results", [])
                st.success("Results:")
                for msg in results:
                    st.write(f"**{msg['role'].capitalize()}**: {msg['content']}")
            else:
                st.error(f"Error: {response.text}")
        except Exception as e:
            st.error(f"Request failed: {e}")
    else:
        st.warning("Please enter a query.")
