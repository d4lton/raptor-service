import asyncio
from fastapi import APIRouter
from services.ExcelService import excel_service_queue

router = APIRouter(prefix="/api/v1/chatgpt", tags=["ChatGPT"])

@router.get("/process_excel/")
async def process_excel():
    print("process_excel create_future")
    future = asyncio.get_event_loop().create_future()
    print("process_excel add to task_queue")
    excel_service_queue.put(future)
    print("process_excel await future")
    result = await future  # Await the result asynchronously
    print("process_excel await done")
    return {"result": result}
