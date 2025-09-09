import os
import time
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from dotenv import load_dotenv
from azure.identity import DefaultAzureCredential
from azure.ai.projects import AIProjectClient
from azure.ai.agents.models import FunctionTool, RequiredFunctionToolCall, ToolOutput, SubmitToolOutputsAction
from user_logic_apps import AzureLogicAppTool, invoke_logicapps_sharepoint

# Load environment variables from .env file
load_dotenv()

app = FastAPI()

class AgentRequest(BaseModel):
    query: str

def run_sharepoint_agent(query: str):
    project_client = AIProjectClient(
        endpoint=os.environ["PROJECT_ENDPOINT"],
        credential=DefaultAzureCredential()
    )
    subscription_id = os.environ["SUBSCRIPTION_ID"]
    resource_group = os.environ["RESOURCE_GROUP"]
    logic_app_name = os.environ["LOGIC_APP_NAME"]
    trigger_name = os.environ["TRIGGER_NAME"]

    logic_app_tool = AzureLogicAppTool(subscription_id, resource_group)
    logic_app_tool.register_logic_app(logic_app_name, trigger_name)

    send_http_request_via_logic_app_sharepoint = invoke_logicapps_sharepoint(logic_app_tool, logic_app_name)
    functions_to_use = {send_http_request_via_logic_app_sharepoint}
    functions = FunctionTool(functions=functions_to_use)

    with project_client:
        agent = project_client.agents.create_agent(
            model=os.environ["MODEL_DEPLOYMENT_NAME"],
            name="Sharepoint Agent",
            instructions="You are an AI Assistant. You must retrieve and reference content stored in Sharepoint folders. "
            "use send_http_request_via_logic_app_sharepoint function to retrieve data from sharepoint about NorthWind Standard Plans.",
            tools=functions.definitions,
        )

        thread = project_client.agents.threads.create()
        message = project_client.agents.messages.create(
            thread_id=thread.id,
            role="user",
            content=query,
        )
        run = project_client.agents.runs.create(thread_id=thread.id, agent_id=agent.id)

        while run.status in ["queued", "in_progress", "requires_action"]:
            time.sleep(1)
            run = project_client.agents.runs.get(thread_id=thread.id, run_id=run.id)

            if run.status == "requires_action" and isinstance(run.required_action, SubmitToolOutputsAction):
                tool_calls = run.required_action.submit_tool_outputs.tool_calls
                if not tool_calls:
                    project_client.agents.runs.cancel(thread_id=thread.id, run_id=run.id)
                    break

                tool_outputs = []
                for tool_call in tool_calls:
                    if isinstance(tool_call, RequiredFunctionToolCall):
                        try:
                            output = functions.execute(tool_call)
                            tool_outputs.append(
                                ToolOutput(
                                    tool_call_id=tool_call.id,
                                    output=output,
                                )
                            )
                        except Exception as e:
                            print(f"Error executing tool_call {tool_call.id}: {e}")
                if tool_outputs:
                    project_client.agents.runs.submit_tool_outputs(thread_id=thread.id, run_id=run.id, tool_outputs=tool_outputs)

        print(f"Run completed with status: {run.status}")
        messages = project_client.agents.messages.list(thread_id=thread.id)
        results = []
        for message in messages:
            results.append({"role": message["role"], "content": message["content"]})
        return results

@app.post("/run-agent")
def run_agent_api(request: AgentRequest):
    try:
        results = run_sharepoint_agent(request.query)
        return {"results": results}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
