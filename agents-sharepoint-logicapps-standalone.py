import os, time
from azure.identity import DefaultAzureCredential
from azure.ai.projects import AIProjectClient
from azure.ai.agents.models import FunctionTool, RequiredFunctionToolCall, ToolOutput, SubmitToolOutputsAction
import json
import datetime
from typing import Any, Callable, Set, Dict, List, Optional
from dotenv import load_dotenv
from user_logic_apps import AzureLogicAppTool, invoke_logicapps_sharepoint

# Load environment variables from .env file
load_dotenv()

# Initialize the AIProjectClient
project_client = AIProjectClient(
    endpoint=os.environ["PROJECT_ENDPOINT"],
    credential=DefaultAzureCredential()
)

# Extract subscription and resource group from environment variables
subscription_id = os.environ["SUBSCRIPTION_ID"]
resource_group = os.environ["RESOURCE_GROUP"]

# Logic App details
logic_app_name = os.environ["LOGIC_APP_NAME"]
trigger_name = os.environ["TRIGGER_NAME"]

# Create and initialize AzureLogicAppTool utility
logic_app_tool = AzureLogicAppTool(subscription_id, resource_group)
logic_app_tool.register_logic_app(logic_app_name, trigger_name)
print(f"Registered logic app '{logic_app_name}' with trigger '{trigger_name}'.")

# Create the function to be used by the agent
send_http_request_via_logic_app_sharepoint = invoke_logicapps_sharepoint(logic_app_tool, logic_app_name)

# result = send_http_request_via_logic_app_sharepoint("NorthWind Standard Plans")
# print(result)

#Prepare the function tools for the agent
functions_to_use = {send_http_request_via_logic_app_sharepoint}
functions = FunctionTool(functions=functions_to_use)

# Create an agent
with project_client:
    # Create an agent with custom functions
    agent = project_client.agents.create_agent(
        model=os.environ["MODEL_DEPLOYMENT_NAME"],
        name="Sharepoint Agent",
        instructions="You are an AI Assistant. You must retrieve and reference content stored in Sharepoint folders. " \
        "use send_http_request_via_logic_app_sharepoint function to retrieve data from sharepoint about NorthWind Standard Plans.",
        tools=functions.definitions,
    )

    print(f"Created agent, ID: {agent.id}")

    # Create a thread for communication
    thread = project_client.agents.threads.create()
    print(f"Created thread, ID: {thread.id}")

    # Send a message to the thread
    message = project_client.agents.messages.create(
        thread_id=thread.id,
        role="user",
        content="Summarize key benefits of NorthWind Standard Plans in a concise manner.",
    )
    print(f"Created message, ID: {message['id']}")

    # Create and process a run for the agent to handle the message
    run = project_client.agents.runs.create(thread_id=thread.id, agent_id=agent.id)
    print(f"Created run, ID: {run.id}")

    # Poll the run status until it is completed or requires action
    while run.status in ["queued", "in_progress", "requires_action"]:
        time.sleep(1)
        run = project_client.agents.runs.get(thread_id=thread.id, run_id=run.id)

        if run.status == "requires_action" and isinstance(run.required_action, SubmitToolOutputsAction):
            tool_calls = run.required_action.submit_tool_outputs.tool_calls
            if not tool_calls:
                print("No tool calls provided - cancelling run")
                project_client.agents.runs.cancel(thread_id=thread.id, run_id=run.id)
                break

            tool_outputs = []
            for tool_call in tool_calls:
                if isinstance(tool_call, RequiredFunctionToolCall):
                    try:
                        print(f"Executing tool call: {tool_call}")
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

    # Fetch and log all messages from the thread
    messages = project_client.agents.messages.list(thread_id=thread.id)
    for message in messages:
        print(f"Role: {message['role']}, Content: {message['content']}")

    # Delete the agent after use
    project_client.agents.delete_agent(agent.id)
    print("Deleted agent")