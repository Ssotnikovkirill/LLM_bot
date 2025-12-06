"""
Telegram бот для сбора требований к проектам с использованием локальной LLM Ollama
Автор: Профессиональный Data Scientist & разработчик
Версия: 1.0.0
"""

import os
import json
import logging
import asyncio
import re
import time
import requests
import aiohttp
import random
from typing import List, Dict, Optional, Tuple, Any, Set
from dataclasses import dataclass, asdict, field
from datetime import datetime, timedelta
from pathlib import Path
from enum import Enum

import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.shapes.autoshape import Shape

from telegram import (
    Update,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    InputFile,
    KeyboardButton,
    ReplyKeyboardMarkup,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
    CallbackQueryHandler,
    CallbackContext,
)

from dotenv import load_dotenv

# Загрузка переменных окружения
load_dotenv()

# ============== КОНФИГУРАЦИЯ ==============
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "tinyllama")

if not TELEGRAM_TOKEN:
    raise RuntimeError("Пожалуйста, установите TELEGRAM_TOKEN в переменных окружения или .env файле")

# Параметры модели
OLLAMA_TIMEOUT = 300  # секунд
OLLAMA_MAX_TOKENS = 1500
OLLAMA_TEMPERATURE = 0.2
OLLAMA_TOP_P = 0.9

# Настройки приложения
SESSIONS_DIR = Path("sessions")
SESSIONS_DIR.mkdir(exist_ok=True)

LOGS_DIR = Path("logs")
LOGS_DIR.mkdir(exist_ok=True)

PP_TEMPLATE_PATH = Path("pres_template.pptx")

# Настройка логирования
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    handlers=[
        logging.FileHandler(LOGS_DIR / "bot.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============== СТАТЕСЫ ДИАЛОГА ==============
class ConversationState(Enum):
    START = 0
    PRIMARY_QUESTIONS = 1
    PRIMARY_CONFIRMATION = 2
    CLARIFYING_QUESTIONS = 3
    PREVIEW = 4
    EDITING = 5
    FINAL = 6
    WAITING_FOR_BP_DETAILS = 7  # Для ввода "Другое" в бизнес-процессе

# ============== МОДЕЛИ ДАННЫХ ==============
@dataclass
class Answer:
    """Ответ на вопрос"""
    question: str
    answer: str
    section: str
    timestamp: datetime = field(default_factory=datetime.now)
    
    def to_dict(self) -> Dict:
        return {
            "question": self.question,
            "answer": self.answer,
            "section": self.section,
            "timestamp": self.timestamp.isoformat()
        }

@dataclass
class Prediction:
    """Предсказание ответа от LLM"""
    question: str
    short_prediction: str
    detailed_prediction: str
    confidence: float = 0.0
    reasoning: str = ""

@dataclass
class OnePagerSections:
    """Разделы одностраничника"""
    project_name: str = ""
    project_goal: str = ""
    business_tasks: str = ""
    as_is: str = ""
    to_be: str = ""
    target_image: str = ""
    used_technologies: str = ""
    data_sources: str = ""
    program_contribution: str = ""
    
    def to_dict(self) -> Dict:
        return asdict(self)

@dataclass
class AnalystBlock:
    """Блок для аналитика"""
    project_name: str = ""
    project_goal: str = ""
    current_situation: str = ""
    future_state: str = ""
    ai_solution: str = ""
    business_value: str = ""
    risks_limitations: str = ""
    data_sources_volume: str = ""
    suitable_technology: str = ""
    current_stage: str = ""
    next_steps: str = ""
    
    def to_dict(self) -> Dict:
        return asdict(self)

@dataclass
class Session:
    """Сессия диалога с пользователем"""
    chat_id: int
    user_id: int
    created_at: datetime = field(default_factory=datetime.now)
    updated_at: datetime = field(default_factory=datetime.now)
    
    # Основные ответы
    primary_answers: List[Answer] = field(default_factory=list)
    current_question_index: int = 0
    
    # Уточняющие вопросы
    clarifying_questions: List[Answer] = field(default_factory=list)
    current_clarifying_index: int = 0
    
    # Предсказания
    current_prediction: Optional[Prediction] = None
    predictions_cache: Dict[str, Prediction] = field(default_factory=dict)
    
    # Результаты
    summary_12: str = ""
    one_pager: OnePagerSections = field(default_factory=OnePagerSections)
    analyst_block: AnalystBlock = field(default_factory=AnalystBlock)
    
    awaiting_q5_q6_response: bool = False
    awaiting_modified_response: bool = False  # Для кнопки "Изменить"
    has_no_problem_in_q5: bool = False  

    # Состояние
    state: ConversationState = ConversationState.START
    awaiting_bp_details: bool = False
    editing_mode: bool = False
    edit_target: Optional[str] = None
    
    # Флаги проверок
    q5_q6_checked: bool = False
    
    def __post_init__(self):
        """Инициализация первичных вопросов"""
        if not self.primary_answers:
            self.primary_answers = [
                Answer(
                    question=q_data["q"],
                    answer="",
                    section=q_data["section"]
                )
                for q_data in PRIMARY_QUESTIONS
            ]
    
    def save(self):
        """Сохранение сессии в файл"""
        session_file = SESSIONS_DIR / f"{self.chat_id}.json"
        data = {
            "chat_id": self.chat_id,
            "user_id": self.user_id,
            "created_at": self.created_at.isoformat(),
            "updated_at": datetime.now().isoformat(),
            "primary_answers": [a.to_dict() for a in self.primary_answers],
            "current_question_index": self.current_question_index,
            "clarifying_questions": [a.to_dict() for a in self.clarifying_questions],
            "current_clarifying_index": self.current_clarifying_index,
            "summary_12": self.summary_12,
            "one_pager": self.one_pager.to_dict(),
            "analyst_block": self.analyst_block.to_dict(),
            "state": self.state.value,
            "awaiting_bp_details": self.awaiting_bp_details,
            "editing_mode": self.editing_mode,
            "edit_target": self.edit_target,
            "q5_q6_checked": self.q5_q6_checked,
            "has_no_problem_in_q5": self.has_no_problem_in_q5,  
            "awaiting_modified_response": self.awaiting_modified_response  
        }
        
        with open(session_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    @classmethod
    def load(cls, chat_id: int) -> Optional['Session']:
        """Загрузка сессии из файла"""
        session_file = SESSIONS_DIR / f"{chat_id}.json"
        
        if not session_file.exists():
            return None
        
        try:
            with open(session_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            session = cls(
                chat_id=data["chat_id"],
                user_id=data.get("user_id", data["chat_id"])
            )
            
            session.created_at = datetime.fromisoformat(data["created_at"])
            session.updated_at = datetime.fromisoformat(data["updated_at"])
            
            # Загрузка основных ответов
            session.primary_answers = [
                Answer(
                    question=item["question"],
                    answer=item["answer"],
                    section=item["section"],
                    timestamp=datetime.fromisoformat(item["timestamp"])
                )
                for item in data.get("primary_answers", [])
            ]
            
            session.current_question_index = data.get("current_question_index", 0)
            
            # Загрузка уточняющих вопросов
            session.clarifying_questions = [
                Answer(
                    question=item["question"],
                    answer=item["answer"],
                    section="Уточняющий вопрос",
                    timestamp=datetime.fromisoformat(item["timestamp"])
                )
                for item in data.get("clarifying_questions", [])
            ]
            
            session.current_clarifying_index = data.get("current_clarifying_index", 0)
            
            # Загрузка результатов
            session.summary_12 = data.get("summary_12", "")
            
            one_pager_data = data.get("one_pager", {})
            session.one_pager = OnePagerSections(**one_pager_data)
            
            analyst_data = data.get("analyst_block", {})
            session.analyst_block = AnalystBlock(**analyst_data)
            
            # Загрузка состояния
            session.state = ConversationState(data.get("state", 0))
            session.awaiting_bp_details = data.get("awaiting_bp_details", False)
            session.editing_mode = data.get("editing_mode", False)
            session.edit_target = data.get("edit_target")
            session.q5_q6_checked = data.get("q5_q6_checked", False)
            session.has_no_problem_in_q5 = data.get("has_no_problem_in_q5", False)  # ДОБАВИТЬ
            session.awaiting_modified_response = data.get("awaiting_modified_response", False)  # ДОБАВИТЬ

            
            return session
            
        except Exception as e:
            logger.error(f"Ошибка загрузки сессии {chat_id}: {e}")
            return None
    
    def reset(self):
        """Сброс сессии"""
        self.primary_answers = [
            Answer(
                question=q_data["q"],
                answer="",
                section=q_data["section"]
            )
            for q_data in PRIMARY_QUESTIONS
        ]
        self.current_question_index = 0
        self.clarifying_questions = []
        self.current_clarifying_index = 0
        self.current_prediction = None
        self.predictions_cache = {}
        self.summary_12 = ""
        self.one_pager = OnePagerSections()
        self.analyst_block = AnalystBlock()
        self.state = ConversationState.START
        self.awaiting_bp_details = False
        self.editing_mode = False
        self.edit_target = None
        self.q5_q6_checked = False
        self.has_no_problem_in_q5 = False  # ДОБАВИТЬ
        self.awaiting_modified_response = False  # ДОБАВИТЬ

# ============== ВОПРОСЫ ==============
PRIMARY_QUESTIONS = [
    {"section": "ОБЩИЕ СВЕДЕНИЯ О ПРОЕКТЕ", "q": "Как коротко называется ваш проект или идея, которую вы хотите реализовать?"},
    {"section": "ОБЩИЕ СВЕДЕНИЯ О ПРОЕКТЕ", "q": "Какой бизнес-процесс затрагивает проект?"},
    {"section": "ОБЩИЕ СВЕДЕНИЯ О ПРОЕКТЕ", "q": "Название бизнес процесса\n(По КТ-001 или его детализация)"},
    {"section": "ТЕКУЩАЯ СИТУАЦИЯ (AS IS)", "q": "Как сейчас решается задача без ИИ?"},
    {"section": "ТЕКУЩАЯ СИТУАЦИЯ (AS IS)", "q": "Какие ключевые проблемы/ограничения вы видите в текущем подходе?"},
    {"section": "ЦЕЛЬ ПРОЕКТА", "q": "Какую ключевую цель вы хотите достичь?\n(сократить затраты/повысить точность/ускорить процессы/автоматизировать принятие решений…)"},
    {"section": "ИСТОЧНИКИ ДАННЫХ", "q": "Какие данные будут на вход?\n(типы, форматы)"},
    {"section": "ИСТОЧНИКИ ДАННЫХ", "q": "Из каких систем получаете данные?\n(назовите конкретные названия, если есть)"},
    {"section": "РЕШАЕМЫЕ БИЗНЕС-ЗАДАЧИ", "q": "Какой ожидаемый результат/выход (формат и пример)?"},
    {"section": "ИНТЕГРАЦИЯ", "q": "Где должно работать решение (назовите системы/платформы/интеграция)?"},
    {"section": "ЦЕЛЕВОЙ ОБРАЗ", "q": "Какой необходим уровень точности системы?\n(в итоговом результате - целевом, и на первом этапе - прототипе)"},
    {"section": "ИНФО О ТЕКУЩЕМ ЭТАПЕ", "q": "На каком этапе сейчас находится проект?\n(идея, анализ, прототип, MVP, пилот, внедрение)?"},
    {"section": "ИНФО О ТЕКУЩЕМ ЭТАПЕ", "q": "Есть ли наработки по этой задаче? Какие?\n(например: прототип)"},
    {"section": "ОРГАНИЗАЦИЯ ПРОЕКТА", "q": "Есть ли у вас готовая команда или требуется найти исполнителя? Внешние подрядчики или всё реализуется внутри компании?"},
]

# ============== КЛАВИАТУРЫ ==============
def get_primary_keyboard(show_start_over: bool = True) -> ReplyKeyboardMarkup:
    """Клавиатура для основных вопросов"""
    buttons = [
        [KeyboardButton("Поясни")],
        [KeyboardButton("Пропустить")]
    ]
    
    if show_start_over:
        buttons.append([KeyboardButton("Начать заново")])
    
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True, one_time_keyboard=True)

def get_business_process_keyboard() -> InlineKeyboardMarkup:
    """Inline клавиатура для выбора бизнес-процесса"""
    buttons = [
        [   
            InlineKeyboardButton("БРД", callback_data="bp_brd"),
            InlineKeyboardButton("БЛПС", callback_data="bp_blps"),
        ],
        [
            InlineKeyboardButton("КФ", callback_data="bp_kf"),
            InlineKeyboardButton("Другое", callback_data="bp_other"),
        ]
    ]
    return InlineKeyboardMarkup(buttons)

def get_prediction_keyboard() -> ReplyKeyboardMarkup:
    """Клавиатура при наличии предсказания"""
    buttons = [
        [KeyboardButton("Принять")],
        [KeyboardButton("Изменить")],
        [KeyboardButton("Ввести свой")],
        [KeyboardButton("Поясни"), KeyboardButton("Пропустить")],
        [KeyboardButton("Начать заново")]
    ]
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True, one_time_keyboard=True)

def get_confirmation_keyboard() -> ReplyKeyboardMarkup:
    """Клавиатура для подтверждения"""
    buttons = [
        [KeyboardButton("Продолжить"), KeyboardButton("Редактировать ответы")],
        [KeyboardButton("Начать заново")]
    ]
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True, one_time_keyboard=True)

def get_q5_q6_keyboard() -> ReplyKeyboardMarkup:
    """Клавиатура для проверки Q5-Q6"""
    buttons = [
        [KeyboardButton("Да, хотим заранее улучшить")],
        [KeyboardButton("Нет, не критично")],
        [KeyboardButton("Пояснить подробнее")]
    ]
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True, one_time_keyboard=True)

def get_final_keyboard() -> ReplyKeyboardMarkup:
    """Клавиатура в конце"""
    buttons = [
        [KeyboardButton("Редактировать результат")],
        [KeyboardButton("Сформировать презентацию")],
        [KeyboardButton("Начать новый проект")]
    ]
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True, one_time_keyboard=True)

# ============== API OLLAMA ==============
class OllamaAPI:
    """Клиент для работы с Ollama API"""
    
    def __init__(self, base_url: str = OLLAMA_BASE_URL):
        self.base_url = base_url.rstrip('/')
        self.session = None
    
    async def __aenter__(self):
        self.session = aiohttp.ClientSession()
        return self
    
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        if self.session:
            await self.session.close()
    
    async def generate(self, prompt: str, system_prompt: str = "", **kwargs) -> Optional[str]:
        """Генерация текста через Ollama"""
        if not self.session:
            self.session = aiohttp.ClientSession()
        
        url = f"{self.base_url}/api/generate"
        
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        
        payload = {
            "model": OLLAMA_MODEL,
            "prompt": prompt,
            "system": system_prompt,
            "stream": False,
            "options": {
                "temperature": kwargs.get("temperature", OLLAMA_TEMPERATURE),
                "top_p": kwargs.get("top_p", OLLAMA_TOP_P),
                "num_predict": kwargs.get("max_tokens", OLLAMA_MAX_TOKENS),
            }
        }
        
        try:
            async with self.session.post(url, json=payload, timeout=aiohttp.ClientTimeout(total=OLLAMA_TIMEOUT)) as response:
                if response.status == 200:
                    result = await response.json()
                    return result.get("response", "").strip()
                else:
                    error_text = await response.text()
                    logger.error(f"Ollama API error {response.status}: {error_text}")
                    return None
        except asyncio.TimeoutError:
            logger.error(f"Ollama API timeout после {OLLAMA_TIMEOUT} секунд")
            return None
        except Exception as e:
            logger.error(f"Ошибка при вызове Ollama API: {e}")
            return None
    
    async def chat(self, messages: List[Dict[str, str]], **kwargs) -> Optional[str]:
        """Чат через Ollama"""
        if not self.session:
            self.session = aiohttp.ClientSession()
        
        url = f"{self.base_url}/api/chat"
        
        payload = {
            "model": OLLAMA_MODEL,
            "messages": messages,
            "stream": False,
            "options": {
                "temperature": kwargs.get("temperature", OLLAMA_TEMPERATURE),
                "top_p": kwargs.get("top_p", OLLAMA_TOP_P),
                "num_predict": kwargs.get("max_tokens", OLLAMA_MAX_TOKENS),
            }
        }
        
        try:
            async with self.session.post(url, json=payload, timeout=aiohttp.ClientTimeout(total=OLLAMA_TIMEOUT)) as response:
                if response.status == 200:
                    result = await response.json()
                    return result.get("message", {}).get("content", "").strip()
                else:
                    error_text = await response.text()
                    logger.error(f"Ollama API error {response.status}: {error_text}")
                    return None
        except asyncio.TimeoutError:
            logger.error(f"Ollama API timeout после {OLLAMA_TIMEOUT} секунд")
            return None
        except Exception as e:
            logger.error(f"Ошибка при вызове Ollama API: {e}")
            return None

# ============== ОБРАБОТЧИКИ LLM ==============
class LLMProcessor:
    """Обработчик LLM запросов"""
    
    def __init__(self):
        self.ollama = OllamaAPI()
        self.cache = {}
    
    async def explain_question(self, session: Session, question: str) -> str:
        """Объяснение вопроса"""
        cache_key = f"explain_{hash(question)}"
        
        if cache_key in self.cache:
            return self.cache[cache_key]
        
        context = ""
        for answer in session.primary_answers[:5]:
            if answer.answer:
                context += f"Вопрос: {answer.question}\nОтвет: {answer.answer}\n\n"
        
        prompt = f"""Ты — деловой аналитик. Объясни коротко (1 предложение), что конкретно нужно указать в ответе на следующий вопрос. 
Не добавляй лишнего, не проси другую информацию. Только поясни, что подразумевает этот вопрос (только этот вопрос). Будь максимально краток, не пиши свои мысли, а пиши только пояснение к вопросу. Строго только на русском языке! Не добавляй лишнего, не проси другую информацию.

Контекст диалога:
{context}

Вопрос для пояснения: {question}

Пояснение (одно короткое предложение):"""
        
        response = await self.ollama.generate(
            prompt=prompt,
            system_prompt="Ты помощник-аналитик, который объясняет вопросы понятным языком. Использует только русский язык, и общается грамотно, без ошибок. Не рассуждай письменно. Пиши четко, коротко, понятно и по делу. Не добавляй лишнего, не проси другую информацию.",
            temperature=0.2,
            max_tokens=80
        )
        
        if response:
            self.cache[cache_key] = response
            return response
        else:
            return "Краткое пояснение временно недоступно."
    
    async def predict_answer(self, session: Session, question: str) -> Optional[Prediction]:
        """Предсказание ответа на вопрос"""
        cache_key = f"predict_{hash(str(session.primary_answers))}_{hash(question)}"
        
        if cache_key in self.cache:
            return self.cache[cache_key]
        
        context = "Контекст диалога:\n"
        for i, answer in enumerate(session.primary_answers):
            if i >= session.current_question_index:
                break
            if answer.answer:
                context += f"Вопрос {i+1}: {answer.question}\nОтвет: {answer.answer}\n\n"
        
        prompt = f"""Ты — опытный профессиональный бизнес-аналитик. Пиши только на русском языке. Не пиши свои рассуждения, будь краток и точен. На основе контекста диалога предскажи, как пользователь может ответить на следующий вопрос. Не добавляй лишнего, не проси другую информацию.
Верни ответ в формате:
SHORT: <короткий вариант ответа (1 предложение)>
DETAILED: <развернутый вариант ответа (3-5 пунктов)>
REASONING: <обоснование, почему ты так решил (2-3 предложения)>

{context}
Следующий вопрос: {question}

Предсказание:"""
        
        response = await self.ollama.generate(
            prompt=prompt,
            system_prompt="Ты профессиональный помощник-аналитик, который детально анализирует контекст диалога и выдает наиболее вероятное предсказание ответа пользователя, как сейчас это у него происходит. Используй только русский язык, пиши грамотно, без ошибок. Не рассуждай письменно. Пиши четко, коротко, понятно и по делу. Не добавляй лишнего, не проси другую информацию.",
            temperature=0.3,
            max_tokens=400
        )
        
        if response:
            prediction = self._parse_prediction(response, question)
            if prediction:
                self.cache[cache_key] = prediction
                return prediction
        
        return None
    
    def _parse_prediction(self, text: str, question: str) -> Optional[Prediction]:
        """Парсинг предсказания из текста"""
        try:
            short_match = re.search(r'SHORT:\s*(.+?)(?:\n|$)', text, re.IGNORECASE)
            detailed_match = re.search(r'DETAILED:\s*(.+?)(?:\nREASONING:|$)', text, re.DOTALL | re.IGNORECASE)
            reasoning_match = re.search(r'REASONING:\s*(.+?)$', text, re.DOTALL | re.IGNORECASE)
            
            short = short_match.group(1).strip() if short_match else ""
            detailed = detailed_match.group(1).strip() if detailed_match else ""
            reasoning = reasoning_match.group(1).strip() if reasoning_match else ""
            
            if short or detailed:
                return Prediction(
                    question=question,
                    short_prediction=short,
                    detailed_prediction=detailed,
                    reasoning=reasoning,
                    confidence=0.7 if detailed else 0.5
                )
        except Exception as e:
            logger.error(f"Ошибка парсинга предсказания: {e}")
        
        return None
    
    async def generate_summary_12(self, session: Session) -> str:
        """Генерация сводки на 12 предложений"""
        context = "Ответы на основные вопросы:\n"
        for answer in session.primary_answers:
            if answer.answer:
                context += f"Вопрос: {answer.question}\nОтвет: {answer.answer}\n\n"
        
        prompt = f"""Ты опытный бизнес системный аналитик, который объясняет вопросы понятным языком. Использует только русский язык, и общается грамотно, без ошибок. Не рассуждай письменно. Пиши четко, коротко, понятно и по делу. На основе ответов сформулируй итоговую сводку, начинающуюся с фразы "Я правильно понял, что..." (в конце этого предложения поставь знак вопроса). На основе всех вопросов и ответов сформулируй деловую проверочную сводку, начинающуюся с фразы "Я правильно понял, что...". В конце этого предложения знак вопроса. Сводка должна быть до 14 предложений (коротких, ясных, аккуратно оформленных по каждому пункту, а не сплошным текстом, чтобы каждый новый раздел был с нового абзаца и по пунктам, а не сплошным текстом - всё четко структурированно и читаемо, по пунктам с разделением на тире в каждом абзаце, вместо перечисления через запятую), в бизнес-формулировках. Не копируй ответы дословно, синтезируй. Используй больше конкретики, названий и терминов, которые были упомянуты в ответах. Также добаь предварительное предположение на основе полученной информации, какую технологию необходимо использовать для реализации этого проекта - подходящую технологию (LLM, RAG, графы, мультиагенты, RL, оптимизация, прикладные ML модели или др.) с детальным обоснованием выбора именно технологии и почему именно эта, а не другая, и какие варианты технологий также можно использовать, но через сравнение выбери лучшую, обосновывая техническим языком с точки зрения Data Science.

Сводка должна быть до 14 предложений, строго на русском языке (допустимо использовать английские сокращения или названия), структурированных по пунктам:
1. Название проекта и суть
2. Затрагиваемый бизнес-процесс
3. Текущая ситуация (AS IS)
4. Ключевые проблемы
5. Цель проекта
6. Источники данных
7. Ожидаемый результат
8. Интеграционные требования
9. Требуемая точность
10. Текущий этап
11. Наличие наработок
12. Организационные вопросы
13. Подходящая технология из области Data Science (например: LLM, RAG, графы, мультиагенты, RL, оптимизация, прикладные ML-модели) с кратким обоснованием выбора.

{context}

Сводка (начни с 'Я правильно понял, что'):"""
        
        response = await self.ollama.generate(
            prompt=prompt,
            system_prompt="Ты опытный бизнес-системный-аналитик, который объясняет вопросы понятным языком. Использует только русский язык, и общается грамотно, без ошибок. Не рассуждай письменно. Пиши четко, коротко, понятно и по делу.",
            temperature=0.2,
            max_tokens=550
        )
        
        if response:
            if not response.startswith("Я правильно понял"):
                response = "Я правильно понял, что " + response
            return response
        else:
            return "Я правильно понял, что информация временно недоступна."
    
    async def generate_clarifying_questions(self, session: Session, max_questions: int = 5) -> List[str]:
        """Генерация уточняющих вопросов"""
        context = "Ответы на основные вопросы:\n"
        for answer in session.primary_answers:
            context += f"Вопрос: {answer.question}\nОтвет: {answer.answer}\n\n"
        
        # Примеры хороших уточняющих вопросов
        examples = """
Примеры хороших уточняющих вопросов:
1. Какие конкретные метрики вы хотите улучшить (например, время обработки документа в часах, процент ошибок)?
2. Можете привести пример типичного документа или процесса, который сейчас обрабатывается вручную?
3. Какие системы используются сейчас и какие данные в них хранятся (названия систем, типы данных)?
4. Какие интеграции с другими системами необходимы (API, базы данных, облачные сервисы)?
5. Какие ограничения по бюджету и срокам существуют?
6. Кто основные пользователи системы и какие у них роли?
7. Какие нормативные требования или стандарты необходимо учитывать?
8. Есть ли исторические данные для обучения моделей и в каком объеме?
9. Какие каналы связи с системой предполагаются (веб-интерфейс, мобильное приложение, API)?
10. Какие риски безопасности данных необходимо учесть?
"""
        
        prompt = f"""Ты — опытный бизнес-аналитик. Общайся строго на русском языке. На основе ответов на основные вопросы сформулируй до {max_questions} уточняющих вопросов, которые помогут лучше понять проект для заполнения одностраничника.

Одностраничник включает разделы (на русском):
- Цель проекта (конкретные формулировки как в примерах, например: Использование больших языковых моделей и интеллектуальных помощников для сопоставления технических НМД и выявления противоречий на основе машинного обучения)
- Решаемые бизнес-задачи (конкретные задачи с цифрами и примерами, например: Согласно приказу 39-П, производится экспертиза СТО ИНТИ экспертами функциональных направлений.  Решение позволит сократить трудозатраты на выполнение ручной сверки документов, оптимизировать процесс экспертизы и выявить противоречия между внешними (отраслевыми, государственными) стандартами/ требованиями с внутренними документами Компании, минимизировать человеческий фактор в выявлении несоответствий. Интеллектуальный помощник охватывает всю цепочку СТО ИНТИ от ГРР до Добычи.)
- AS IS (текущая ситуация с конкретными проблемами и цифрами, например: Более 300 документов ежегодно проходят сверку и согласование профильных служб: 1.Ручной процесс экспертизы и выявления несоответствий, противоречий; 2.Необходимость повторной сверки/ перепроверки по многим направлениям; 3.Колоссальные трудозатраты на выполнение экспертизы внешней документации.)
- TO BE (будущее состояние с использованием ИИ, конкретные улучшения, например: На основе онтологической модели и сформированного алгоритма выполняется автоматическая сверка НМД с выявлением отличий и противоречий. Выявленные противоречия/отклонения от внутренних НМД рассматривает эксперт и принимает решение о корректном варианте формулировок/ уточняет реестр замечаний. В решении также присутствует автоматическая повторная сверка с ранее выданными замечаниями при прохождении второго круга экспертизы документа.)
- Целевой образ и результат проекта (конкретные ожидаемые результаты, например: Цифровой инструмент с интерактивным интерфейсом и функционалом, направленный на: 1. формирование реестра  замечаний к СТО ИНТИ в части соблюдения ГОСТ по предметной области, выявления расхождений логического описания требований документа, заполнение установленного чек-листа экспертизы; 2. автоматическую подготовку реестра выявленных противоречий между проверяемым документом и стандартами, регламентами Компании; 3.ускорение процесса экспертизы документации за счет автоматизации процесса.)
- Организационный объем проекта (например: Реализованная онтология и алгоритм будет разработан по всем техническим направлениям разведки и добычи для использования специалистами ГПН НТЦ, ГЕО, ГПН ТП, по мере масштабирования  - всеми ДО.)
- Используемые технологии (конкретные технологии с обоснованием, например: LLM, NLP, RAG, Онтологическая модель.)
- Источники данных (конкретные системы и типы данных, например: ЦНМД, ИССК)
- Вклад в программу (конкретные микросервисы или модули, которые будут полезные для других проектов в будущем, например: Микросервис на базе LLM, который позволяет находить противоречия между внешними отраслевыми стандартами и внутренними НМД.)

{examples}

{context}

Сформулируй {max_questions} конкретных, четких уточняющих вопросов на русском языке (только вопросы, каждый с новой строки с номером):"""
        
        response = await self.ollama.generate(
            prompt=prompt,
            system_prompt="Ты аналитик, который задает конкретные, четкие уточняющие вопросы для сбора требований. Строго на русском языке. Учитывай весь контекст диалога, предыдущие вопросы и ответы. Не рассуждай письменно, а пиши четко и по делу.",
            temperature=0.4,
            max_tokens=400
        )
        
        if response:
            questions = []
            lines = response.strip().split('\n')
            for line in lines:
                # Убираем номера и маркеры
                clean_line = re.sub(r'^\d+[\.\)]\s*', '', line.strip())
                clean_line = re.sub(r'^[-\*•]\s*', '', clean_line)
                if clean_line and len(clean_line) > 10 and '?' in clean_line:
                    questions.append(clean_line)
            
            # Удаляем дубликаты
            seen = set()
            unique_questions = []
            for q in questions:
                q_lower = q.lower()
                if q_lower not in seen and len(q_lower.split()) > 3:
                    seen.add(q_lower)
                    unique_questions.append(q)
            
            return unique_questions[:max_questions]
        
        return []
    

#     async def generate_one_pager(self, session: Session) -> OnePagerSections:
#         """Генерация разделов одностраничника"""
#         context = "Ответы на вопросы:\n"
        
#         # Основные ответы
#         for answer in session.primary_answers:
#             if answer.answer:
#                 context += f"{answer.section}: {answer.question}\nОтвет: {answer.answer}\n\n"
        
#         # Уточняющие ответы
#         if session.clarifying_questions:
#             context += "Уточняющие ответы:\n"
#             for answer in session.clarifying_questions:
#                 if answer.answer:
#                     context += f"Вопрос: {answer.question}\nОтвет: {answer.answer}\n\n"
        
#         prompt = f"""Ты — опытный продуктовый аналитик. На основе информации о проекте заполни разделы одностраничника строго по шаблонам.

# Контекст проекта:
# {context}

# Заполни каждый раздел ОДНОСТРАНИЧНИКА:

# 1. ЦЕЛЬ ПРОЕКТА (формулировка как в примерах):
# Примеры:
# - "Использование больших языковых моделей и интеллектуальных помощников для сопоставления технических НМД и выявления противоречий на основе машинного обучения"
# - "Разработка интеллектуальной системы на основе LLM для автоматизации поиска информации в базе технической документации и анализа инцидентов на газовом оборудовании"
# - "Разработка инструмента для снижения сроков концептуального проектирования и повышения качества разрабатываемых технических решений при выборе низкотемпературной технологии подготовки газа с использованием мультиагентных технологий"

# 2. РЕШАЕМЫЕ БИЗНЕС-ЗАДАЧИ (конкретные задачи с эффектом):
# Примеры:
# - "Согласно приказу 39-П, производится экспертиза СТО ИНТИ экспертами функциональных направлений. Решение позволит сократить трудозатраты на выполнение ручной сверки документов, оптимизировать процесс экспертизы и выявить противоречия между внешними (отраслевыми, государственными) стандартами/требованиями с внутренними документами Компании, минимизировать человеческий фактор в выявлении несоответствий. Интеллектуальный помощник охватывает всю цепочку СТО ИНТИ от ГРР до Добычи."
# - "Внедрение ИИ на базе LLM, основанной на архитектуре трансформеров (GPT), позволит автоматизировать обработку технической документации, классифицировать инциденты по типам оборудования (ГПА, КУ, турбодетандеры) и формировать рекомендации, минимизируя незапланированные остановы, снижая затраты на обслуживание."

# 3. AS IS (текущая ситуация с конкретными проблемами):
# Примеры:
# - "Более 300 документов ежегодно проходят сверку и согласование профильных служб: 1. Ручной процесс экспертизы и выявления несоответствий, противоречий, 2. Необходимость повторной сверки/перепроверки по многим направлениям, 3. Колоссальные трудозатраты на выполнение экспертизы внешней документации"
# - "Сотрудники вручную осуществляют поиск в разрозненных архивах (локальные файлы, бумажные документы) и тратят до нескольких часов на анализ инцидентов, что приводит к длительным простоям оборудования, росту операционных затрат из-за ошибок в интерпретации данных и отсутствия механизма применения исторических данных для data-driven оптимизации регламентов"

# 4. TO BE (будущее состояние с ИИ):
# Примеры:
# - "На основе онтологической модели и сформированного алгоритма выполняется автоматическая сверка НМД с выявлением отличий и противоречий. Выявленные противоречия/отклонения от внутренних НМД рассматривает эксперт и принимает решение о корректном варианте формулировок/уточняет реестр замечаний. В решении также присутствует автоматическая повторная сверка с ранее выданными замечаниями при прохождении второго круга экспертизы документа."
# - "Внедрение AI-системы с семантическим поиском по технической документации (технические регламенты, журналы ТОиР) и автоматизированным анализом инцидентов на базе LLM позволит сократить время на поиск информации до 1-2 минут, генерировать рекомендации по устранению неполадок со ссылками на нормативы, прогнозировать риски и снижать простои за счет превентивных мер, обеспечивая безопасность и соответствие требованиям при работе с газовым оборудованием."

# 5. ЦЕЛЕВОЙ ОБРАЗ И РЕЗУЛЬТАТ ПРОЕКТА:
# Примеры:
# - "Цифровой инструмент с интерактивным интерфейсом и функционалом, направленный на: 1. формирование реестра замечаний к СТО ИНТИ в части соблюдения ГОСТ по предметной области, выявления расхождений логического описания требований документа, заполнение установленного чек-листа экспертизы; 2. автоматическую подготовку реестра выявленных противоречий между проверяемым документом и стандартами, регламентами Компании; 3. ускорение процесса экспертизы документации за счет автоматизации процесса."
# - "Интеллектуальная платформа на основе технологий ИИ, которая объединяет данные из различных источников в единую экспертно-аналитическую систему, обеспечивая персонал точными инструкциями для устранения неполадок, снижая аварийность за счет перехода от ручного управления знаниями к data-driven стратегиям. Качественный эффект: 1. Переход от ручного анализа к «интеллектуальной» автоматизации, 2. Снижение зависимости от экспертного опыта, 3. Повышение скорость адаптации к изменениям"

# 6. ИСПОЛЬЗУЕМЫЕ ТЕХНОЛОГИИ (конкретные технологии с обоснованием выбора):
# Примеры:
# - "LLM, NLP, RAG, Онтологическая модель"
# - "LLM на архитектуре GPT, RAG, OCR"
# - "Системный инжиниринг, Интеллектуальные мультиагентные технологии, Суррогатное моделирование (прокси-модели на основе ИИ), Большие языковые модели для анализа базы данных оборудования, Продвинутые методы оптимизации"

# 7. ИСТОЧНИКИ ДАННЫХ (конкретные системы):
# Примеры:
# - "ЦНМД, ИССК"
# - "Проектные команды (данные ГиР, региональный обзор, константы проектирования)"

# 8. ВКЛАД В ПРОГРАММУ (конкретные микросервисы/модули):
# Примеры:
# - "Микросервис на базе LLM, который позволяет находить противоречия между внешними отраслевыми стандартами и внутренними НМД."
# - "Микролит, на базе суррогатной модели AspenTech Hysys по низкотемпературной технологии подготовки газа для прогнозирования результатов расчета."

# Верни ответ в строгом JSON формате:
# {{
#   "project_goal": "текст цели проекта",
#   "business_tasks": "текст бизнес-задач",
#   "as_is": "текст AS IS",
#   "to_be": "текст TO BE",
#   "target_image": "текст целевого образа",
#   "used_technologies": "текст технологий",
#   "data_sources": "текст источников данных",
#   "program_contribution": "текст вклада в программу"
# }}"""
        
#         response = await self.ollama.generate(
#             prompt=prompt,
#             system_prompt="Ты аналитик, который заполняет одностраничник строго по шаблонам.",
#             temperature=0.2,
#             max_tokens=1500
#         )
        
#         if response:
#             try:
#                 # Ищем JSON в ответе
#                 json_match = re.search(r'\{.*\}', response, re.DOTALL)
#                 if json_match:
#                     data = json.loads(json_match.group())
#                     return OnePagerSections(
#                         project_name=session.primary_answers[0].answer if session.primary_answers else "",
#                         project_goal=data.get("project_goal", ""),
#                         business_tasks=data.get("business_tasks", ""),
#                         as_is=data.get("as_is", ""),
#                         to_be=data.get("to_be", ""),
#                         target_image=data.get("target_image", ""),
#                         used_technologies=data.get("used_technologies", ""),
#                         data_sources=data.get("data_sources", ""),
#                         program_contribution=data.get("program_contribution", "")
#                     )
#             except json.JSONDecodeError as e:
#                 logger.error(f"Ошибка парсинга JSON одностраничника: {e}")
        
#         # Fallback
#         return OnePagerSections(
#             project_name=session.primary_answers[0].answer if session.primary_answers else "",
#             project_goal="Цель проекта",
#             business_tasks="Бизнес-задачи",
#             as_is="Текущая ситуация",
#             to_be="Будущее состояние",
#             target_image="Целевой образ",
#             used_technologies="Технологии",
#             data_sources="Источники данных",
#             program_contribution="Вклад в программу"
#         )
    
    async def generate_one_pager(self, session: Session) -> OnePagerSections:
        """Генерация разделов одностраничника с улучшенными промптами"""
        context = "Ответы на вопросы:\n\n"
        
        # Основные ответы
        for i, answer in enumerate(session.primary_answers):
            if answer.answer and answer.answer != "<пропущено>":
                context += f"{i+1}. {answer.question}\nОтвет: {answer.answer}\n\n"
        
        # Уточняющие ответы
        if session.clarifying_questions:
            context += "Уточняющие ответы:\n"
            for answer in session.clarifying_questions:
                if answer.answer and answer.answer != "<пропущено>":
                    context += f"Вопрос: {answer.question}\nОтвет: {answer.answer}\n\n"
        
        # Улучшенный промпт с более четкими инструкциями
        prompt = f"""Ты — опытный продуктовый аналитик в нефтегазовой отрасли. На основе информации о проекте заполни разделы одностраничника строго по шаблонам-примерам.

    КОНТЕКСТ ПРОЕКТА:
    {context}

    ЗАПОЛНИ КАЖДЫЙ РАЗДЕЛ ОДНОСТРАНИЧНИКА СТРОГО ПО ШАБЛОНАМ (но подробно, не одним словом, а как в шаблонах-примерах, строго на русском языке, но допустимо использовать английские сокращения):

    1. ЦЕЛЬ ПРОЕКТА (формулировка как в примерах):
    Примеры:
    - "Использование больших языковых моделей и интеллектуальных помощников для сопоставления технических НМД и выявления противоречий на основе машинного обучения"
    - "Разработка интеллектуальной системы на основе LLM для автоматизации поиска информации в базе технической документации и анализа инцидентов на газовом оборудовании"
    - "Разработка инструмента для снижения сроков концептуального проектирования и повышения качества разрабатываемых технических решений при выборе низкотемпературной технологии подготовки газа с использованием мультиагентных технологий"

    Твоя формулировка цели (начинай с глагола):

    2. РЕШАЕМЫЕ БИЗНЕС-ЗАДАЧИ (конкретные задачи с эффектом):
    Примеры:
    - "Согласно приказу 39-П, производится экспертиза СТО ИНТИ экспертами функциональных направлений. Решение позволит сократить трудозатраты на выполнение ручной сверки документов, оптимизировать процесс экспертизы и выявить противоречия между внешними (отраслевыми, государственными) стандартами/требованиями с внутренними документами Компании, минимизировать человеческий фактор в выявлении несоответствий. Интеллектуальный помощник охватывает всю цепочку СТО ИНТИ от ГРР до Добычи."
    - "Внедрение ИИ на базе LLМ, основанной на архитектуре трансформеров (GPT), позволит автоматизировать обработку технической документации, классифицировать инциденты по типам оборудования (ГПА, КУ, турбодетандеры) и формировать рекомендации, минимизируя незапланированные остановы, снижая затраты на обслуживание."

    Твои формулировки бизнес-задач (3-5 конкретных задач):

    3. AS IS (текущая ситуация с конкретными проблемами):
    Примеры:
    - "Более 300 документов ежегодно проходят сверку и согласование профильных служб: 1. Ручной процесс экспертизы и выявления несоответствий, противоречий, 2. Необходимость повторной сверки/перепроверки по многим направлениям, 3. Колоссальные трудозатраты на выполнение экспертизы внешней документации"
    - "Сотрудники вручную осуществляют поиск в разрозненных архивах (локальные файлы, бумажные документы) и тратят до нескольких часов на анализ инцидентов, что приводит к длительным простоям оборудования, росту операционных затрат из-за ошибок в интерпретации данных и отсутствия механизма применения исторических данных для data-driven оптимизации регламентов"

    Твое описание AS IS (с цифрами и конкретными проблемами):

    4. TO BE (будущее состояние с ИИ):
    Примеры:
    - "На основе онтологической модели и сформированного алгоритма выполняется автоматическая сверка НМД с выявлением отличий и противоречий. Выявленные противоречия/отклонения от внутренних НМД рассматривает эксперт и принимает решение о корректном варианте формулировок/уточняет реестр замечаний. В решении также присутствует автоматическая повторная сверка с ранее выданными замечаниями при прохождении второго круга экспертизы документа."
    - "Внедрение AI-системы с семантическим поиском по технической документации (технические регламенты, журналы ТОиР) и автоматизированным анализом инцидентов на базе LLM позволит сократить время на поиск информации до 1-2 минут, генерировать рекомендации по устранению неполадок со ссылками на нормативы, прогнозировать риски и снижать простои за счет превентивных мер, обеспечивая безопасность и соответствие требованиям при работе с газовым оборудованием."

    Твое описание TO BE (с конкретными улучшениями):

    5. ЦЕЛЕВОЙ ОБРАЗ И РЕЗУЛЬТАТ ПРОЕКТА:
    Примеры:
    - "Цифровой инструмент с интерактивным интерфейсом и функционалом, направленный на: 1. формирование реестра замечаний к СТО ИНТИ в части соблюдения ГОСТ по предметной области, выявления расхождений логического описания требований документа, заполнение установленного чек-листа экспертизы; 2. автоматическую подготовку реестра выявленных противоречий между проверяемым документом и стандартами, регламентами Компании; 3. ускорение процесса экспертизы документации за счет автоматизации процесса."
    - "Интеллектуальная платформа на основе технологий ИИ, которая объединяет данные из различных источников в единую экспертно-аналитическую систему, обеспечивая персонал точными инструкции для устранения неполадок, снижая аварийность за счет перехода от ручного управления знаниями к data-driven стратегиям. Качественный эффект: 1. Переход от ручного анализа к «интеллектуальной» автоматизации, 2. Снижение зависимости от экспертного опыта, 3. Повышение скорость адаптации к изменениям"

    Твой целевой образ (с конкретными функциями):

    6. ОРГАНИЗАЦИОННЫЙ ОБЪЁМ ПРОЕКТА:
    Примеры:
    - "Реализованная онтология и алгоритм будет разработан по всем техническим направлениям разведки и добычи для использования специалистами ГПН НТЦ, ГЕО, ГПН ТП, по мере масштабирования  - всеми ДО"

    Твой организационный объем проекта (строго по примеру):

    7. ИСПОЛЬЗУЕМЫЕ ТЕХНОЛОГИИ (конкретные технологии с обоснованием выбора):
    Примеры:
    - "LLM, NLP, RAG, Онтологическая модель"
    - "LLM на архитектуре GPT, RAG, OCR"
    - "Системный инжиниринг, Интеллектуальные мультиагентные технологии, Суррогатное моделирование (прокси-модели на основе ИИ), Большие языковые модели для анализа базы данных оборудования, Продвинутые методы оптимизации"

    Твои технологии (перечисли через запятую):

    8. ИСТОЧНИКИ ДАННЫХ (конкретные системы):
    Примеры:
    - "ЦНМД, ИССК"
    - "Проектные команды (данные ГиР, региональный обзор, константы проектирования)"

    Твои источники данных (перечисли через запятую):

    9. ВКЛАД В ПРОГРАММУ (конкретные микросервисы/модули):
    Примеры:
    - "Микросервис на базе LLM, который позволяет находить противоречия между внешними отраслевыми стандартами и внутренними НМД."
    - "Микролит, на базе суррогатной модели AspenTech Hysys по низкотемпературной технологии подготовки газа для прогнозирования результатов расчета."

    Твой вклад в программу (1-2 конкретных микросервиса):

    ВЕРНИ ОТВЕТ ТОЛЬКО В JSON ФОРМАТЕ БЕЗ КАКИХ-ЛИБО ПРЕДВАРИТЕЛЬНЫХ ТЕКСТОВ:
    {{
    "project_goal": "твоя формулировка цели",
    "business_tasks": "твои бизнес-задачи",
    "as_is": "твое описание AS IS",
    "to_be": "твое описание TO BE",
    "target_image": "твой целевой образ",
    "project scope": "твой организационный объем проекта",
    "used_technologies": "твои технологии",
    "data_sources": "твои источники данных",
    "program_contribution": "твой вклад в программу"
    }}"""
        
        response = await self.ollama.generate(
            prompt=prompt,
            system_prompt="Ты профессиональный русский бизнес-системный аналитик, который заполняет одностраничник строго по шаблонам, развернуто, логично, идеально всё обдумав. Возвращай ТОЛЬКО JSON.",
            temperature=0.1,  # Низкая температура для более точного следования шаблонам
            max_tokens=500
        )
        
        if response:
            try:
                # Ищем JSON в ответе
                json_match = re.search(r'\{.*\}', response, re.DOTALL)
                if json_match:
                    data = json.loads(json_match.group())
                    return OnePagerSections(
                        project_name=session.primary_answers[0].answer if session.primary_answers and session.primary_answers[0].answer else "",
                        project_goal=data.get("project_goal", "Цель проекта не сгенерирована"),
                        business_tasks=data.get("business_tasks", "Бизнес-задачи не сгенерированы"),
                        as_is=data.get("as_is", "AS IS не сгенерировано"),
                        to_be=data.get("to_be", "TO BE не сгенерировано"),
                        target_image=data.get("target_image", "Целевой образ не сгенерирован"),
                        used_technologies=data.get("used_technologies", "Технологии не сгенерированы"),
                        data_sources=data.get("data_sources", "Источники данных не сгенерированы"),
                        program_contribution=data.get("program_contribution", "Вклад в программу не сгенерирован")
                    )
            except json.JSONDecodeError as e:
                logger.error(f"Ошибка парсинга JSON одностраничника: {e}")
                logger.error(f"Ответ LLM: {response}")
        
        # Fallback
        return OnePagerSections(
            project_name=session.primary_answers[0].answer if session.primary_answers else "",
            project_goal="Цель проекта не сгенерирована",
            business_tasks="Бизнес-задачи не сгенерированы",
            as_is="AS IS не сгенерировано",
            to_be="TO BE не сгенерировано",
            target_image="Целевой образ не сгенерирован",
            used_technologies="Технологии не сгенерированы",
            data_sources="Источники данных не сгенерированы",
            program_contribution="Вклад в программу не сгенерирован"
        )

    async def generate_analyst_block(self, session: Session) -> AnalystBlock:
        """Генерация блока для аналитика"""
        context = "Ответы на вопросы:\n"
        
        for answer in session.primary_answers:
            if answer.answer:
                context += f"{answer.section}: {answer.question}\nОтвет: {answer.answer}\n\n"
        
        if session.clarifying_questions:
            context += "Уточняющие ответы:\n"
            for answer in session.clarifying_questions:
                if answer.answer:
                    context += f"Вопрос: {answer.question}\nОтвет: {answer.answer}\n\n"
        
        prompt = f"""Ты профессиональный бизнес-системный аналитик. На основе информации о проекте заполни блок для аналитика. Строго на русском, правильно и логично, без ошибок, анализируя и используя весь контекст диалога, все ответы на вопросы. Синтерзируй. Используй больше конкретики. Пиши преимущественно на русском языке, допустимо использовать английские сокращения в названиях.

Контекст проекта:
{context}

Заполни КАЖДЫЙ раздел блока для аналитика:

1. НАЗВАНИЕ ПРОЕКТА: короткое название
2. ЦЕЛЬ ПРОЕКТА: четкая формулировка цели
3. ТЕКУЩАЯ СИТУАЦИЯ (AS IS): описание текущего состояния с проблемами
4. БУДУЩЕЕ СОСТОЯНИЕ (TO BE): описание целевого состояния с ИИ
5. РЕШЕНИЕ С ИСПОЛЬЗОВАНИЕМ ИИ: как именно ИИ решит проблему
6. КЛЮЧЕВАЯ ЦЕННОСТЬ ДЛЯ БИЗНЕСА: бизнес-ценность (экономия, эффективность)
7. ОСНОВНЫЕ РИСКИ И ОГРАНИЧЕНИЯ: потенциальные риски и ограничения
8. ИСТОЧНИКИ И ОБЪЕМ ДАННЫХ: какие данные и в каком объеме
9. ПОДХОДЯЩАЯ ТЕХНОЛОГИЯ (с обоснованием выбора): конкретная технология с детальным техническим обоснованием почему именно она, а не другие, но можно написать комбинации технологий (напрмиер: LLM и RAG)
10. ТЕКУЩИЙ ЭТАП И СЛЕДУЮЩИЕ ШАГИ: этап проекта и план действий

Технологии для выбора (с обоснованием):
- LLM (Large Language Models) - для обработки текста, диалоговых систем
- RAG (Retrieval-Augmented Generation) - для поиска по документам с контекстом
- Графовые базы данных - для сложных связей и онтологий
- Мультиагентные системы - для распределенных задач и переговоров
- Reinforcement Learning - для оптимизации и адаптивных систем
- Прикладные ML-модели - для специфических задач (классификация, регрессия)
- Оптимизационные алгоритмы - для поиска оптимальных решений

Верни ответ в строгом JSON формате:
{{
  "project_name": "название",
  "project_goal": "цель",
  "current_situation": "AS IS",
  "future_state": "TO BE",
  "ai_solution": "решение с ИИ",
  "business_value": "ценность",
  "risks_limitations": "риски",
  "data_sources_volume": "данные",
  "suitable_technology": "технология с обоснованием",
  "current_stage": "этап",
  "next_steps": "шаги"
}}"""
        
        response = await self.ollama.generate(
            prompt=prompt,
            system_prompt="Ты профессионвльный системный-бизнес аналитик, создающий детальные блоки для аналитиков на русском языке. Делай всё четко, правильно, логично, без ошибок. Анализируй полностью весь предыдущий диалог, все ответы на вопросы и общий контекст.",
            temperature=0.2,
            max_tokens=500
        )
        
        if response:
            try:
                json_match = re.search(r'\{.*\}', response, re.DOTALL)
                if json_match:
                    data = json.loads(json_match.group())
                    return AnalystBlock(**data)
            except json.JSONDecodeError as e:
                logger.error(f"Ошибка парсинга JSON блока аналитика: {e}")
        
        # Fallback
        return AnalystBlock(
            project_name=session.primary_answers[0].answer if session.primary_answers else "",
            project_goal="Цель проекта",
            current_situation="Текущая ситуация",
            future_state="Будущее состояние",
            ai_solution="Решение с ИИ",
            business_value="Бизнес-ценность",
            risks_limitations="Риски и ограничения",
            data_sources_volume="Источники данных",
            suitable_technology="Подходящая технология",
            current_stage="Текущий этап",
            next_steps="Следующие шаги"
        )

# ============== МЕНЕДЖЕР СЕССИЙ ==============
class SessionManager:

    """Менеджер сессий"""
    
    def __init__(self):
        self.sessions: Dict[int, Session] = {}
        self.llm = LLMProcessor()
    
    def get_session(self, chat_id: int, user_id: int) -> Session:
        """Получение или создание сессии"""
        if chat_id not in self.sessions:
            loaded = Session.load(chat_id)
            if loaded:
                self.sessions[chat_id] = loaded
            else:
                self.sessions[chat_id] = Session(chat_id=chat_id, user_id=user_id)
        
        return self.sessions[chat_id]
    
    def save_session(self, session: Session):
        """Сохранение сессии"""
        session.save()
        self.sessions[session.chat_id] = session

# ============== ПРЕЗЕНТАЦИЯ ==============
# class PowerPointGenerator:
#     """Генератор презентации PowerPoint"""
    
#     def __init__(self, template_path: Path = PP_TEMPLATE_PATH):
#         self.template_path = template_path
        
#         if not template_path.exists():
#             logger.warning(f"Шаблон презентации не найден: {template_path}")
#             self.template = Presentation()  # Создаем пустую презентацию
#             self.has_template = False
#         else:
#             self.template = Presentation(template_path)
#             self.has_template = True
    
#     def generate(self, one_pager: OnePagerSections, session: Session) -> Path:
#         """Генерация презентации"""
#         if self.has_template:
#             prs = Presentation(self.template_path)
#         else:
#             prs = Presentation()
#             # Создаем титульный слайд
#             title_slide_layout = prs.slide_layouts[0]
#             slide = prs.slides.add_slide(title_slide_layout)
#             title = slide.shapes.title
#             subtitle = slide.placeholders[1]
            
#             title.text = one_pager.project_name or "Проект"
#             subtitle.text = "Одностраничник проекта"
        
#         # Сохраняем
#         output_path = SESSIONS_DIR / f"onepager_{session.chat_id}.pptx"
#         prs.save(output_path)
        
#         # Создаем текстовый файл с содержимым для слайда
#         content = self._create_slide_content(one_pager)
#         txt_path = SESSIONS_DIR / f"onepager_{session.chat_id}_content.txt"
#         with open(txt_path, 'w', encoding='utf-8') as f:
#             f.write(content)
        
#         logger.info(f"Создана презентация: {output_path}")
#         return output_path
    
#     def _create_slide_content(self, one_pager: OnePagerSections) -> str:
#         """Создание содержимого для слайда"""
#         content = f"""Одностраничник проекта: {one_pager.project_name}

# ЦЕЛЬ ПРОЕКТА:
# {one_pager.project_goal}

# РЕШАЕМЫЕ БИЗНЕС-ЗАДАЧИ:
# {one_pager.business_tasks}

# AS IS (ТЕКУЩАЯ СИТУАЦИЯ):
# {one_pager.as_is}

# TO BE (БУДУЩЕЕ СОСТОЯНИЕ):
# {one_pager.to_be}

# ЦЕЛЕВОЙ ОБРАЗ И РЕЗУЛЬТАТ ПРОЕКТА:
# {one_pager.target_image}

# ИСПОЛЬЗУЕМЫЕ ТЕХНОЛОГИИ:
# {one_pager.used_technologies}

# ИСТОЧНИКИ ДАННЫХ:
# {one_pager.data_sources}

# ВКЛАД В ПРОГРАММУ:
# {one_pager.program_contribution}
# """
#         return content

class PowerPointGenerator:
    """Генератор презентации PowerPoint"""
    
    def __init__(self, template_path: Path = PP_TEMPLATE_PATH):
        self.template_path = template_path
        
        if not template_path.exists():
            logger.warning(f"Шаблон презентации не найден: {template_path}")
            self.template = Presentation()  # Создаем пустую презентацию
            self.has_template = False
        else:
            self.template = Presentation(template_path)
            self.has_template = True
    
    def generate(self, one_pager: OnePagerSections, session: Session) -> Path:
        """Генерация презентации с заполнением всех разделов"""
        if self.has_template:
            prs = Presentation(self.template_path)
        else:
            prs = Presentation()
            
            # Создаем слайд для каждого раздела
            self._create_slides(prs, one_pager)
        
        # Сохраняем
        output_path = SESSIONS_DIR / f"onepager_{session.chat_id}.pptx"
        prs.save(output_path)
        
        logger.info(f"Создана презентация: {output_path}")
        return output_path
    
    def _create_slides(self, prs: Presentation, one_pager: OnePagerSections):
        """Создание слайдов с данными"""
        
        # Слайд 1: Титульный
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        title.text = one_pager.project_name or "Проект"
        subtitle.text = "Одностраничник проекта"
        
        # Слайд 2: Цель проекта
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "ЦЕЛЬ ПРОЕКТА"
        content.text = one_pager.project_goal or "Цель проекта не указана"
        
        # Слайд 3: Бизнес-задачи
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "РЕШАЕМЫЕ БИЗНЕС-ЗАДАЧИ"
        content.text = one_pager.business_tasks or "Бизнес-задачи не указаны"
        
        # Слайд 4: AS IS
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "AS IS (ТЕКУЩАЯ СИТУАЦИЯ)"
        content.text = one_pager.as_is or "Текущая ситуация не описана"
        
        # Слайд 5: TO BE
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "TO BE (БУДУЩЕЕ СОСТОЯНИЕ)"
        content.text = one_pager.to_be or "Будущее состояние не описано"
        
        # Слайд 6: Целевой образ
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "ЦЕЛЕВОЙ ОБРАЗ И РЕЗУЛЬТАТ"
        content.text = one_pager.target_image or "Целевой образ не описан"
        
        # Слайд 7: Технологии
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "ИСПОЛЬЗУЕМЫЕ ТЕХНОЛОГИИ"
        content.text = one_pager.used_technologies or "Технологии не указаны"
        
        # Слайд 8: Источники данных
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "ИСТОЧНИКИ ДАННЫХ"
        content.text = one_pager.data_sources or "Источники данных не указаны"
        
        # Слайд 9: Вклад в программу
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = "ВКЛАД В ПРОГРАММУ"
        content.text = one_pager.program_contribution or "Вклад в программу не указан"

# ============== ОСНОВНОЙ БОТ ==============
class RequirementsBot:
    """Основной класс бота"""
    
    def __init__(self):
        self.session_manager = SessionManager()
        self.powerpoint_gen = PowerPointGenerator()
        self.llm = LLMProcessor()
        
    async def start(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработчик команды /start"""
        chat_id = update.effective_chat.id
        user_id = update.effective_user.id
        
        session = self.session_manager.get_session(chat_id, user_id)
        session.state = ConversationState.START
        session.reset()
        self.session_manager.save_session(session)
        
        welcome_text = """
👋 Добро пожаловать в бот для сбора требований к проекту!

Я помогу вам:
1. Собрать информацию о вашем проекте
2. Уточнить детали через умные вопросы
3. Сформировать структурированный одностраничник
4. Определить подходящие технологии ИИ
5. Сформировать презентацию

Команды:
/fill - начать сбор требований
/export - экспорт результатов
/cancel - отмена текущей сессии
/help - помощь

Начнем сбор требований? Нажмите /fill или кнопку ниже.
"""
        
        keyboard = [[KeyboardButton("/fill")]]
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
        
        await update.message.reply_text(welcome_text, reply_markup=reply_markup)
        
        return ConversationState.START.value
    
    async def fill_start(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Начало сбора требований"""
        chat_id = update.effective_chat.id
        user_id = update.effective_user.id
        
        session = self.session_manager.get_session(chat_id, user_id)
        session.reset()
        session.state = ConversationState.PRIMARY_QUESTIONS
        self.session_manager.save_session(session)
        
        intro_text = """
📋 Начинаем сбор основных требований.

После каждого вопроса доступны кнопки:
• Поясни - объяснение вопроса
• Пропустить - пропустить вопрос
• Начать заново - начать с начала

Я буду анализировать ваши ответы и предлагать варианты для следующих вопросов. Вы можете принять мои предложения или ввести свой вариант.

Готовы? Поехали! 🚀
"""
        
        await update.message.reply_text(intro_text, reply_markup=ReplyKeyboardRemove())
        
        # Запускаем фоновую задачу для предсказания первого вопроса
        # asyncio.create_task(self._generate_prediction_background(session))
        
        return await self.ask_current_question(update, context)
    
    async def ask_current_question(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Задаем текущий вопрос"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        # Проверяем, все ли вопросы заданы
        if session.current_question_index >= len(session.primary_answers):
            return await self.finish_primary_questions(update, context)
        
        current_answer = session.primary_answers[session.current_question_index]
        question_text = current_answer.question
        section_text = current_answer.section

            # ПРОВЕРКА: для первого вопроса (индекс 0) не показываем предсказание
        if session.current_question_index == 0:
            # Для первого вопроса показываем обычный вопрос без предсказания
            formatted_question = f"【{section_text}】\n\n{question_text}"
            
            await update.message.reply_text(
                formatted_question,
                reply_markup=get_primary_keyboard()
            )
            return ConversationState.PRIMARY_QUESTIONS.value
        
        # Проверка Q5-Q6
        # if session.current_question_index == 5 and not session.q5_q6_checked:  # Q6 (индекс 5)
        #     await self.check_q5_q6(session, update)
        #     return ConversationState.PRIMARY_QUESTIONS.value
        
        # Форматирование вопроса
        formatted_question = f"【{section_text}】\n\n{question_text}"
        
        # Для второго вопроса - inline клавиатура
        if session.current_question_index == 1:  # Второй вопрос (индекс 1)
            keyboard = get_business_process_keyboard()
            await update.message.reply_text(formatted_question, reply_markup=keyboard)
            return ConversationState.PRIMARY_QUESTIONS.value
        
        question_key = f"predict_{session.chat_id}_{session.current_question_index}"

        if question_key not in session.predictions_cache:
            await update.message.reply_text("🤖 Думаю...")
            prediction = await self.llm.predict_answer(session, question_text)

            if prediction:
                session.predictions_cache[question_key] = prediction
                session.current_prediction = prediction
                self.session_manager.save_session(session)



        
        # Проверяем наличие предсказания
        if session.current_prediction and session.current_prediction.question == question_text:
            # Показываем предсказание
            prediction_text = self._format_prediction(session.current_prediction)
            full_text = f"{formatted_question}\n\n{prediction_text}"
            
            await update.message.reply_text(
                full_text,
                reply_markup=get_prediction_keyboard()
            )
        else:
            # Обычный вопрос
            await update.message.reply_text(
                formatted_question,
                reply_markup=get_primary_keyboard()
            )
            
            # Запускаем фоновую задачу для предсказания следующего вопроса
            asyncio.create_task(self._generate_prediction_background(session))
        
        return ConversationState.PRIMARY_QUESTIONS.value
    
    def _format_prediction(self, prediction: Prediction) -> str:
        """Форматирование предсказания для отображения"""
        text = "💡 *Предположение бота:*\n\n"
        text += f"*Коротко:* {prediction.short_prediction}\n\n"
        
        if prediction.detailed_prediction:
            text += "*Подробно:*\n"
            # Разбиваем на пункты
            lines = prediction.detailed_prediction.split('\n')
            for line in lines:
                if line.strip():
                    text += f"• {line.strip()}\n"
        
        if prediction.reasoning:
            text += f"\n*Обоснование:* {prediction.reasoning}\n"
        
        text += "\nВыберите действие: Принять / Изменить / Ввести свой"
        return text
    
    async def _generate_prediction_background(self, session: Session):
        """Фоновая генерация предсказания"""
        try:
                # НЕ генерируем предсказание для первого вопроса (индекс 0)
            if session.current_question_index == 0:
                return
            
            if session.current_question_index >= len(session.primary_answers):
                return
            
            current_answer = session.primary_answers[session.current_question_index]
            question_text = current_answer.question
            
            # Проверяем кэш
            cache_key = f"predict_{session.chat_id}_{session.current_question_index}"
            if cache_key in session.predictions_cache:
                session.current_prediction = session.predictions_cache[cache_key]
                session.save()
                return
            
            # Генерируем предсказание
            prediction = await self.llm.predict_answer(session, question_text)
            
            if prediction:
                session.current_prediction = prediction
                session.predictions_cache[cache_key] = prediction
                session.save()
                logger.info(f"Сгенерировано предсказание для вопроса {session.current_question_index}")
        except Exception as e:
            logger.error(f"Ошибка в фоновой генерации предсказания: {e}")
    
    async def handle_business_process_callback(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработчик выбора бизнес-процесса"""
        query = update.callback_query
        await query.answer()
        
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        if query.data == "bp_other":
            # Пользователь выбрал "Другое"
            session.awaiting_bp_details = True
            session.state = ConversationState.WAITING_FOR_BP_DETAILS
            self.session_manager.save_session(session)
            
            await query.edit_message_text("Пожалуйста, введите название вашего бизнес-процесса:")
            return ConversationState.WAITING_FOR_BP_DETAILS.value
        
        if query.data == "bp_brd":
            # Пользователь выбрал "Другое"
            session.awaiting_bp_details = True
            session.state = ConversationState.WAITING_FOR_BP_DETAILS
            self.session_manager.save_session(session)
            
            await query.edit_message_text("Пожалуйста, введите название вашего бизнес-процесса с клавиатуры:\n(напомню: у вас БРД)")
            return ConversationState.WAITING_FOR_BP_DETAILS.value
        
        if query.data == "bp_blps":
            # Пользователь выбрал "Другое"
            session.awaiting_bp_details = True
            session.state = ConversationState.WAITING_FOR_BP_DETAILS
            self.session_manager.save_session(session)
            
            await query.edit_message_text("Пожалуйста, введите название вашего бизнес-процесса с клавиатуры:\n(напомню: у вас БЛПС)")
            return ConversationState.WAITING_FOR_BP_DETAILS.value
        
        if query.data == "bp_kf":
            # Пользователь выбрал "Другое"
            session.awaiting_bp_details = True
            session.state = ConversationState.WAITING_FOR_BP_DETAILS
            self.session_manager.save_session(session)
            
            await query.edit_message_text("Пожалуйста, введите название вашего бизнес-процесса с клавиатуры:\n(напомню: у вас КФ)")
            return ConversationState.WAITING_FOR_BP_DETAILS.value
        
        # Обработка выбора из списка
        bp_mapping = {
            # "bp_brd": "БРД (блок разведки и добычи)",
            # "bp_blps": "БЛПС (блок логистики/переработки/сбыта)",
            # "bp_kf": "КФ (корп. функция)"
        }
        
        selected_bp = bp_mapping.get(query.data, query.data)
        
        # Сохраняем ответ
        session.primary_answers[1].answer = selected_bp
        session.current_question_index += 1
        session.awaiting_bp_details = False
        self.session_manager.save_session(session)
        
        await query.edit_message_text(f"Выбрано: {selected_bp}")
        
        # Переходим к следующему вопросу
        return await self.ask_current_question(update, context)


    async def handle_primary_answer(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
            
        """Обработка ответа на основной вопрос"""
        chat_id = update.effective_chat.id
        user_id = update.effective_user.id
        text = update.message.text.strip()
        
        session = self.session_manager.get_session(chat_id, user_id)
        
        # 1. Проверка, если ожидается измененный ответ после нажатия "Изменить"
        if hasattr(session, 'awaiting_modified_response') and session.awaiting_modified_response:
            session.primary_answers[session.current_question_index].answer = text
            session.awaiting_modified_response = False
            session.current_question_index += 1
            session.current_prediction = None
            self.session_manager.save_session(session)
            
            await update.message.reply_text("✅ Ответ изменен.")
            return await self.ask_current_question(update, context)
        
        # 2. Обработка специальных команд
        if text == "Поясни":
            return await self.handle_explain(update, context)
        elif text == "Пропустить":
            return await self.handle_skip(update, context)
        elif text == "Начать заново":
            return await self.fill_start(update, context)
        # Для первого вопроса не показываем кнопки предсказания
        elif session.current_question_index == 0:
            # Для первого вопроса просто обрабатываем ответ
            if session.current_question_index < len(session.primary_answers):
                session.primary_answers[session.current_question_index].answer = text
                session.current_question_index += 1
                session.current_prediction = None
                self.session_manager.save_session(session)
                
                return await self.ask_current_question(update, context)
        elif text == "Принять":
            return await self.handle_accept_prediction(update, context)
        elif text == "Изменить":
            return await self.handle_modify_prediction(update, context)
        elif text == "Ввести свой":
            await update.message.reply_text("Пожалуйста, введите ваш вариант ответа:")
            return ConversationState.PRIMARY_QUESTIONS.value
        
        # 3. Обработка ответа на проверку Q5-Q6 (если мы в режиме проверки)
        if session.awaiting_q5_q6_response:
            return await self.handle_q5_q6_response(update, context)
        
        # 4. Обычный текстовый ответ (для вопросов после первого)
        if session.current_question_index < len(session.primary_answers):
            # Сохраняем ответ
            session.primary_answers[session.current_question_index].answer = text
            
            # Проверка для Q5 (индекс 4): проверяем ТОЛЬКО Q5
            if session.current_question_index == 4:  # Q5 (индекс 4)
                q5_answer = text.lower()
                
                # Ключевые слова для определения отсутствия проблем
                no_problem_keywords = ["нет", "отсутств", "не наблюда", "не вижу", "нет проблем", 
                                    "всё нормально", "устраивает", "не имеется", "проблем нет"]
                
                has_no_problem = any(keyword in q5_answer for keyword in no_problem_keywords)
                
                # Если в Q5 указано, что проблем нет, устанавливаем флаг
                if has_no_problem:
                    session.has_no_problem_in_q5 = True
                else:
                    session.has_no_problem_in_q5 = False
            
            # Проверка для Q6 (индекс 5): теперь проверяем ОБА ответа
            elif session.current_question_index == 5:  # Q6 (индекс 5)
                q6_answer = text.lower()
                
                # Ключевые слова для целей улучшения
                improvement_keywords = ["сократ", "оптимиз", "повыс", "автоматиз", "ускор", 
                                    "точн", "улучш", "эффектив", "улучшить"]
                
                has_improvement_goal = any(keyword in q6_answer for keyword in improvement_keywords)
                
                # Если в Q5 были указаны "нет проблем" И в Q6 есть цели улучшения
                if getattr(session, 'has_no_problem_in_q5', False) and has_improvement_goal and not session.q5_q6_checked:
                    # Запускаем проверку
                    await self.check_q5_q6(session, update)
                    return ConversationState.PRIMARY_QUESTIONS.value
            
            # Увеличиваем индекс вопроса
            session.current_question_index += 1
            session.current_prediction = None  # Сбрасываем предсказание
            self.session_manager.save_session(session)
            
            # Переходим к следующему вопросу
            return await self.ask_current_question(update, context)
        
        await update.message.reply_text("Пожалуйста, используйте кнопки для навигации.")
        return ConversationState.PRIMARY_QUESTIONS.value

    
    async def handle_explain(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка команды 'Поясни'"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        if session.current_question_index >= len(session.primary_answers):
            await update.message.reply_text("Все вопросы уже заданы.")
            return ConversationState.PRIMARY_QUESTIONS.value
        
        current_question = session.primary_answers[session.current_question_index].question
        
        # Запрашиваем объяснение у LLM
        explanation = await self.llm.explain_question(session, current_question)
        
        response = f"📝 *Пояснение к вопросу:*\n\n{explanation}\n\nПожалуйста, введите ответ:"
        
        await update.message.reply_text(response, reply_markup=get_primary_keyboard())
        
        return ConversationState.PRIMARY_QUESTIONS.value
    
    async def handle_skip(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка команды 'Пропустить'"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        if session.current_question_index < len(session.primary_answers):
            session.primary_answers[session.current_question_index].answer = "<пропущено>"
            session.current_question_index += 1
            session.current_prediction = None
            self.session_manager.save_session(session)
            
            return await self.ask_current_question(update, context)
        
        return ConversationState.PRIMARY_QUESTIONS.value
    
    async def handle_accept_prediction(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Принятие предсказания"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        if session.current_prediction and session.current_question_index < len(session.primary_answers):
            # Используем короткое предсказание как ответ
            session.primary_answers[session.current_question_index].answer = session.current_prediction.short_prediction
            session.current_question_index += 1
            session.current_prediction = None
            self.session_manager.save_session(session)
            
            await update.message.reply_text("✅ Ответ принят. Спасибо!")
            
            return await self.ask_current_question(update, context)
        
        await update.message.reply_text("Нет доступного предсказания для принятия.")
        return ConversationState.PRIMARY_QUESTIONS.value
    
    
    async def handle_modify_prediction(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Изменение предсказания"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        # Устанавливаем флаг, что мы ожидаем ввод от пользователя
        session.awaiting_modified_response = True
        self.session_manager.save_session(session)
        
        await update.message.reply_text(
            "Пожалуйста, введите исправленный вариант ответа:",
            reply_markup=ReplyKeyboardRemove()
        )
        # Оставляем в том же состоянии, чтобы принять ответ
        return ConversationState.PRIMARY_QUESTIONS.value
    
    async def check_q5_q6(self, session: Session, update: Update):
        """Проверка Q5-Q6 на противоречия"""
        q5_answer = session.primary_answers[4].answer.lower() if session.primary_answers[4].answer else ""
        q6_answer = session.primary_answers[5].answer.lower() if len(session.primary_answers) > 5 else ""
        
        # Ключевые слова для определения отсутствия проблем
        no_problem_keywords = ["нет", "отсутств", "не наблюда", "не вижу", "нет проблем", "всё нормально", "устраивает"]
        
        # Ключевые слова для целей улучшения
        improvement_keywords = ["сократ", "оптимиз", "повыс", "автоматиз", "ускор", "точн", "улучш", "эффектив"]
        
        has_no_problem = any(keyword in q5_answer for keyword in no_problem_keywords)
        has_improvement_goal = any(keyword in q6_answer for keyword in improvement_keywords)
        
        if has_no_problem and has_improvement_goal:
            question_text = """
🤔 Я заметил потенциальное противоречие:

Вы указали, что явных проблем нет (или минимальны), но при этом хотите оптимизировать/улучшить процесс.

Пожалуйста, уточните:
1. Действительно ли проблем нет, но вы хотите улучшить процесс заранее?
2. Или проблемы есть, но вы их не указали явно?

Это поможет правильно сформулировать цель проекта.
"""
            
            await update.message.reply_text(question_text, reply_markup=get_q5_q6_keyboard())
            session.q5_q6_checked = True
            self.session_manager.save_session(session)
    
    async def handle_q5_q6_response(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка ответа на проверку Q5-Q6"""
        text = update.message.text.strip()
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        if "да" in text.lower() or "заранее" in text.lower():
            # Пользователь подтверждает улучшение без явных проблем
            if len(session.primary_answers) > 5:
                current_answer = session.primary_answers[5].answer
                session.primary_answers[5].answer = f"{current_answer} (проактивное улучшение)"
            
            await update.message.reply_text(
                "✅ Понял. Отмечаю как проактивное улучшение процесса.",
                reply_markup=get_primary_keyboard()
            )
            
        elif "нет" in text.lower() or "не критично" in text.lower():
            # Пользователь говорит, что улучшение не критично
            await update.message.reply_text(
                "✅ Понял. Учитываю как некритичное улучшение.",
                reply_markup=get_primary_keyboard()
            )
        
        elif "пояснить" in text.lower():
            # Запрос дополнительного пояснения
            explanation = """
Проактивное улучшение — это когда мы улучшаем процесс, даже если явных проблем нет.
Например:
• Автоматизация рутинных задач для высвобождения времени сотрудников
• Внедрение новых технологий для будущей масштабируемости
• Улучшение качества данных для будущих аналитических задач

Если у вас есть такие цели — это нормально и правильно!
"""
            await update.message.reply_text(explanation, reply_markup=get_q5_q6_keyboard())
            return ConversationState.PRIMARY_QUESTIONS.value
        
        # Продолжаем со следующим вопросом
        session.q5_q6_checked = True
        self.session_manager.save_session(session)
        
        return await self.ask_current_question(update, context)
    
    async def finish_primary_questions(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Завершение основных вопросов"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        # Генерируем сводку на 12 предложений
        await update.message.reply_text("⏳ Формирую сводку на основе ваших ответов...")
        
        summary = await self.llm.generate_summary_12(session)
        session.summary_12 = summary
        session.state = ConversationState.PRIMARY_CONFIRMATION
        self.session_manager.save_session(session)
        
        # Сохраняем сводку в файл
        summary_file = SESSIONS_DIR / f"summary_{chat_id}.md"
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write(f"# Сводка проекта\n\n{summary}\n\n")
            f.write("## Ответы на основные вопросы:\n")
            for i, answer in enumerate(session.primary_answers):
                f.write(f"{i+1}. {answer.question}\n   Ответ: {answer.answer}\n\n")
        
        # Отправляем сводку пользователю
        response = f"""
📊 *Краткая сводка-уточнение:*

{summary}

---
*Что дальше?*
1. *Продолжить* — перейти к уточняющим вопросам для детализации
2. *Редактировать ответы* — внести изменения в ответы на основные вопросы
3. *Начать заново* — начать сбор требований заново
"""
        
        await update.message.reply_text(response, reply_markup=get_confirmation_keyboard())
        
        return ConversationState.PRIMARY_CONFIRMATION.value
    
    async def handle_confirmation(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка подтверждения после основных вопросов"""
        text = update.message.text.strip()
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        if text == "Продолжить":
            # Генерируем уточняющие вопросы
            await update.message.reply_text("⏳ Формирую уточняющие вопросы для детализации...")
            
            clarifying_questions = await self.llm.generate_clarifying_questions(session, max_questions=5)
            
            if not clarifying_questions:
                await update.message.reply_text("ℹ️ Дополнительные уточнения не требуются. Формирую итоговые документы...")
                return await self.generate_final_documents(update, context)
            
            # Сохраняем уточняющие вопросы
            session.clarifying_questions = [
                Answer(question=q, answer="", section="Уточняющий вопрос")
                for q in clarifying_questions
            ]
            session.current_clarifying_index = 0
            session.state = ConversationState.CLARIFYING_QUESTIONS
            self.session_manager.save_session(session)
            
            # Задаем первый уточняющий вопрос
            first_question = session.clarifying_questions[0].question
            question_text = f"🔍 *Уточняющий вопрос 1/{len(clarifying_questions)}:*\n\n{first_question}"
            
            await update.message.reply_text(question_text, reply_markup=get_primary_keyboard())
            
            return ConversationState.CLARIFYING_QUESTIONS.value
        
        elif text == "Редактировать ответы":
            # Переход в режим редактирования
            return await self.start_editing(update, context)
        
        elif text == "Начать заново":
            return await self.fill_start(update, context)
        
        return ConversationState.PRIMARY_CONFIRMATION.value
    
    async def start_editing(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Начало режима редактирования"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        # Формируем список вопросов для редактирования
        questions_list = ""
        for i, answer in enumerate(session.primary_answers):
            status = "✅" if answer.answer and answer.answer != "<пропущено>" else "⏳"
            questions_list += f"{i+1}. {status} {answer.question[:50]}...\n"
        
        instructions = """
✏️ *Режим редактирования*

Для редактирования ответа:
1. Напишите номер вопроса (1-14)
2. Или напишите фразу с ключевыми словами, например:
   - "измени название проекта"
   - "исправь цель"
   - "редактирую вопрос 3"

Для выхода из режима редактирования: /continue
"""
        
        await update.message.reply_text(
            f"{instructions}\n\n{questions_list}",
            reply_markup=ReplyKeyboardRemove()
        )
        
        session.state = ConversationState.EDITING
        session.editing_mode = True
        self.session_manager.save_session(session)
        
        return ConversationState.EDITING.value
    
    async def handle_editing(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка команд редактирования"""
        text = update.message.text.strip()
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        # Команда продолжения
        if text.lower() == "/continue":
            session.editing_mode = False
            session.state = ConversationState.PRIMARY_CONFIRMATION
            self.session_manager.save_session(session)
            
            await update.message.reply_text(
                "Редактирование завершено.",
                reply_markup=get_confirmation_keyboard()
            )
            return ConversationState.PRIMARY_CONFIRMATION.value
        
        # Поиск номера вопроса
        numbers = re.findall(r'\d+', text)
        if numbers:
            question_num = int(numbers[0]) - 1  # Индексация с 0
            
            if 0 <= question_num < len(session.primary_answers):
                session.current_question_index = question_num
                session.edit_target = "primary"
                self.session_manager.save_session(session)
                
                question_text = session.primary_answers[question_num].question
                current_answer = session.primary_answers[question_num].answer
                
                response = f"""
Редактирование вопроса {question_num + 1}:

*Вопрос:* {question_text}

*Текущий ответ:* {current_answer if current_answer else "Нет ответа"}

Пожалуйста, введите новый ответ:
"""
                await update.message.reply_text(response)
                return ConversationState.EDITING.value
        
        # Поиск по ключевым словам
        for i, answer in enumerate(session.primary_answers):
            question_lower = answer.question.lower()
            text_lower = text.lower()
            
            # Проверяем совпадение ключевых слов
            keywords = text_lower.split()
            matches = sum(1 for keyword in keywords if keyword in question_lower and len(keyword) > 3)
            
            if matches > 0:
                session.current_question_index = i
                session.edit_target = "primary"
                self.session_manager.save_session(session)
                
                response = f"""
Нашел совпадение с вопросом {i + 1}:

*Вопрос:* {answer.question}

*Текущий ответ:* {answer.answer if answer.answer else "Нет ответа"}

Пожалуйста, введите новый ответ:
"""
                await update.message.reply_text(response)
                return ConversationState.EDITING.value
        
        # Если ничего не найдено
        await update.message.reply_text(
            "Не удалось найти вопрос для редактирования. Попробуйте указать номер (1-14) или более конкретные ключевые слова."
        )
        return ConversationState.EDITING.value
    
    async def handle_edit_response(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка нового ответа в режиме редактирования"""
        text = update.message.text.strip()
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        if session.edit_target == "primary" and session.current_question_index < len(session.primary_answers):
            # Обновляем ответ
            session.primary_answers[session.current_question_index].answer = text
            session.edit_target = None
            
            await update.message.reply_text(f"✅ Ответ на вопрос {session.current_question_index + 1} обновлен.")
            
            # Возвращаемся к списку вопросов
            return await self.start_editing(update, context)
        
        elif session.edit_target == "clarifying" and session.current_question_index < len(session.clarifying_questions):
            session.clarifying_questions[session.current_question_index].answer = text
            session.edit_target = None
        
            await update.message.reply_text(f"✅ Ответ на вопрос {session.current_question_index + 1} обновлен.")
        else:
            await update.message.reply_text("Ошибка редактирования. Попробуйте еще раз.")

        return ConversationState.EDITING.value
    
    async def handle_clarifying_answer(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка ответа на уточняющий вопрос"""
        text = update.message.text.strip()
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        # Обработка команд
        if text == "Поясни":
            current_question = session.clarifying_questions[session.current_clarifying_index].question
            explanation = await self.llm.explain_question(session, current_question)
            
            await update.message.reply_text(
                f"📝 *Пояснение:*\n\n{explanation}\n\nПожалуйста, введите ответ:",
                reply_markup=get_primary_keyboard()
            )
            return ConversationState.CLARIFYING_QUESTIONS.value
        
        elif text == "Пропустить":
            session.clarifying_questions[session.current_clarifying_index].answer = "<пропущено>"
            session.current_clarifying_index += 1
        
        elif text == "Начать заново":
            return await self.fill_start(update, context)
        
        else:
            # Обычный ответ
            session.clarifying_questions[session.current_clarifying_index].answer = text
            session.current_clarifying_index += 1
        
        self.session_manager.save_session(session)
        
        # Проверяем, все ли вопросы заданы
        if session.current_clarifying_index >= len(session.clarifying_questions):
            # Все уточняющие вопросы заданы
            await update.message.reply_text("✅ Все уточняющие вопросы получены. Формирую итоговые документы...")
            return await self.generate_final_documents(update, context)
        
        # Задаем следующий вопрос
        next_question = session.clarifying_questions[session.current_clarifying_index].question
        question_text = f"🔍 *Уточняющий вопрос {session.current_clarifying_index + 1}/{len(session.clarifying_questions)}:*\n\n{next_question}"
        
        await update.message.reply_text(question_text, reply_markup=get_primary_keyboard())
        
        return ConversationState.CLARIFYING_QUESTIONS.value
    
    async def generate_final_documents(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Генерация итоговых документов"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        await update.message.reply_text("⏳ Формирую итоговые документы... Это займет несколько минут.")
        
        # Генерируем одностраничник
        one_pager = await self.llm.generate_one_pager(session)
        session.one_pager = one_pager
        
        # Генерируем блок для аналитика
        analyst_block = await self.llm.generate_analyst_block(session)
        session.analyst_block = analyst_block
        
        session.state = ConversationState.FINAL
        self.session_manager.save_session(session)
        
        # Сохраняем полные данные
        self._save_full_data(session)
        
        # Отправляем результаты
        await self.send_results(update, context, session)
        
        return ConversationState.FINAL.value
    
    def _save_full_data(self, session: Session):
        """Сохранение полных данных"""
        # JSON с полными данными
        full_data = {
            "chat_id": session.chat_id,
            "user_id": session.user_id,
            "created_at": session.created_at.isoformat(),
            "updated_at": datetime.now().isoformat(),
            "summary_12": session.summary_12,
            "one_pager": session.one_pager.to_dict(),
            "analyst_block": session.analyst_block.to_dict(),
            "primary_answers": [a.to_dict() for a in session.primary_answers],
            "clarifying_answers": [a.to_dict() for a in session.clarifying_questions]
        }
        
        json_path = SESSIONS_DIR / f"full_data_{session.chat_id}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(full_data, f, ensure_ascii=False, indent=2)
        
        # Markdown с форматированием
        md_content = self._create_markdown_content(session)
        md_path = SESSIONS_DIR / f"onepager_{session.chat_id}.md"
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(md_content)
        
        logger.info(f"Сохранены данные для сессии {session.chat_id}")

    
    async def edit_name_command(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда редактирования названия проекта"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        session.editing_mode = True
        session.state = ConversationState.EDITING
        session.current_question_index = 0  # Первый вопрос - название
        session.edit_target = "primary"
        self.session_manager.save_session(session)
        
        question_text = session.primary_answers[0].question
        current_answer = session.primary_answers[0].answer
        
        await update.message.reply_text(
            f"✏️ Редактирование названия проекта:\n\n"
            f"*Текущий ответ:* {current_answer if current_answer else 'Нет ответа'}\n\n"
            f"Пожалуйста, введите новое название:",
            parse_mode='Markdown'
        )
        
        return ConversationState.EDITING.value

    async def edit_goal_command(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда редактирования цели проекта"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        session.editing_mode = True
        session.state = ConversationState.EDITING
        session.current_question_index = 5  # Шестой вопрос - цель
        session.edit_target = "primary"
        self.session_manager.save_session(session)
        
        question_text = session.primary_answers[5].question
        current_answer = session.primary_answers[5].answer
        
        await update.message.reply_text(
            f"✏️ Редактирование цели проекта:\n\n"
            f"*Вопрос:* {question_text}\n\n"
            f"*Текущий ответ:* {current_answer if current_answer else 'Нет ответа'}\n\n"
            f"Пожалуйста, введите новую цель:",
            parse_mode='Markdown'
        )
        
        return ConversationState.EDITING.value

    async def edit_all_command(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда полного редактирования"""
        return await self.start_editing(update, context)

    async def _handle_back_to_primary(self, update: Update, context: ContextTypes.DEFAULT_TYPE, session: Session = None, text: str = None):
        """Обработка возврата к первостепенным вопросам"""
        chat_id = update.effective_chat.id
        if not session:
            session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        session.state = ConversationState.PRIMARY_CONFIRMATION
        session.editing_mode = False
        self.session_manager.save_session(session)
        
        await update.message.reply_text(
            "Возвращаюсь к первостепенным вопросам.",
            reply_markup=get_confirmation_keyboard()
        )
        
        return ConversationState.PRIMARY_CONFIRMATION.value

    
    def _create_markdown_content(self, session: Session) -> str:
        """Создание Markdown контента"""
        md = f"""# Одностраничник проекта: {session.one_pager.project_name}

## 📊 Сводка на 12 предложений

{session.summary_12}

---

## 🎯 РАЗДЕЛЫ ОДНОСТРАНИЧНИКА

### 1. ЦЕЛЬ ПРОЕКТА
{session.one_pager.project_goal}

### 2. РЕШАЕМЫЕ БИЗНЕС-ЗАДАЧИ
{session.one_pager.business_tasks}

### 3. AS IS (ТЕКУЩАЯ СИТУАЦИЯ)
{session.one_pager.as_is}

### 4. TO BE (БУДУЩЕЕ СОСТОЯНИЕ)
{session.one_pager.to_be}

### 5. ЦЕЛЕВОЙ ОБРАЗ И РЕЗУЛЬТАТ ПРОЕКТА
{session.one_pager.target_image}

### 6. ИСПОЛЬЗУЕМЫЕ ТЕХНОЛОГИИ
{session.one_pager.used_technologies}

### 7. ИСТОЧНИКИ ДАННЫХ
{session.one_pager.data_sources}

### 8. ВКЛАД В ПРОГРАММУ
{session.one_pager.program_contribution}

---

## 📈 БЛОК ДЛЯ АНАЛИТИКА

### 1. НАЗВАНИЕ ПРОЕКТА
{session.analyst_block.project_name}

### 2. ЦЕЛЬ ПРОЕКТА
{session.analyst_block.project_goal}

### 3. ТЕКУЩАЯ СИТУАЦИЯ (AS IS)
{session.analyst_block.current_situation}

### 4. БУДУЩЕЕ СОСТОЯНИЕ (TO BE)
{session.analyst_block.future_state}

### 5. РЕШЕНИЕ С ИСПОЛЬЗОВАНИЕМ ИИ
{session.analyst_block.ai_solution}

### 6. КЛЮЧЕВАЯ ЦЕННОСТЬ ДЛЯ БИЗНЕСА
{session.analyst_block.business_value}

### 7. ОСНОВНЫЕ РИСКИ И ОГРАНИЧЕНИЯ
{session.analyst_block.risks_limitations}

### 8. ИСТОЧНИКИ И ОБЪЕМ ДАННЫХ
{session.analyst_block.data_sources_volume}

### 9. ПОДХОДЯЩАЯ ТЕХНОЛОГИЯ
{session.analyst_block.suitable_technology}

### 10. ТЕКУЩИЙ ЭТАП И СЛЕДУЮЩИЕ ШАГИ
{session.analyst_block.current_stage}

---

## 📝 ИСТОРИЯ ДИАЛОГА

### Основные вопросы и ответы:
"""
        
        for i, answer in enumerate(session.primary_answers):
            md += f"\n**{i+1}. {answer.question}**\n"
            md += f"Ответ: {answer.answer if answer.answer else 'Нет ответа'}\n"
        
        if session.clarifying_questions:
            md += "\n### Уточняющие вопросы и ответы:\n"
            for i, answer in enumerate(session.clarifying_questions):
                md += f"\n**Уточнение {i+1}. {answer.question}**\n"
                md += f"Ответ: {answer.answer if answer.answer else 'Нет ответа'}\n"
        
        md += f"\n---\n*Сгенерировано: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n"
        
        return md
    
    async def send_results(self, update: Update, context: ContextTypes.DEFAULT_TYPE, session: Session):
        """Отправка результатов пользователю"""
        # Отправляем разделы одностраничника
        onepager_text = f"""
🎯 *ОДНОСТРАНИЧНИК ПРОЕКТА*

*Проект:* {session.one_pager.project_name}

*1. ЦЕЛЬ ПРОЕКТА*
{session.one_pager.project_goal}

*2. РЕШАЕМЫЕ БИЗНЕС-ЗАДАЧИ*
{session.one_pager.business_tasks}

*3. AS IS (ТЕКУЩАЯ СИТУАЦИЯ)*
{session.one_pager.as_is}

*4. TO BE (БУДУЩЕЕ СОСТОЯНИЕ)*
{session.one_pager.to_be}

*5. ЦЕЛЕВОЙ ОБРАЗ И РЕЗУЛЬТАТ ПРОЕКТА*
{session.one_pager.target_image}

*6. ИСПОЛЬЗУЕМЫЕ ТЕХНОЛОГИИ*
{session.one_pager.used_technologies}

*7. ИСТОЧНИКИ ДАННЫХ*
{session.one_pager.data_sources}

*8. ВКЛАД В ПРОГРАММУ*
{session.one_pager.program_contribution}
"""
        
        await update.message.reply_text(onepager_text)
        
        # Отправляем блок для аналитика
        analyst_text = f"""
📈 *БЛОК ДЛЯ АНАЛИТИКА*

*1. НАЗВАНИЕ ПРОЕКТА*
{session.analyst_block.project_name}

*2. ЦЕЛЬ ПРОЕКТА*
{session.analyst_block.project_goal}

*3. ТЕКУЩАЯ СИТУАЦИЯ (AS IS)*
{session.analyst_block.current_situation}

*4. БУДУЩЕЕ СОСТОЯНИЕ (TO BE)*
{session.analyst_block.future_state}

*5. РЕШЕНИЕ С ИСПОЛЬЗОВАНИЕМ ИИ*
{session.analyst_block.ai_solution}

*6. КЛЮЧЕВАЯ ЦЕННОСТЬ ДЛЯ БИЗНЕСА*
{session.analyst_block.business_value}

*7. ОСНОВНЫЕ РИСКИ И ОГРАНИЧЕНИЯ*
{session.analyst_block.risks_limitations}

*8. ИСТОЧНИКИ И ОБЪЕМ ДАННЫХ*
{session.analyst_block.data_sources_volume}

*9. ПОДХОДЯЩАЯ ТЕХНОЛОГИЯ*
{session.analyst_block.suitable_technology}

*10. ТЕКУЩИЙ ЭТАП И СЛЕДУЮЩИЕ ШАГИ*
{session.analyst_block.current_stage}
"""
        
        await update.message.reply_text(analyst_text)
        
        # Отправляем файлы
        json_path = SESSIONS_DIR / f"full_data_{session.chat_id}.json"
        md_path = SESSIONS_DIR / f"onepager_{session.chat_id}.md"
        
        try:
            # Отправляем Markdown
            with open(md_path, 'rb') as f:
                await update.message.reply_document(
                    document=InputFile(f, filename=f"Одностраничник_{session.one_pager.project_name}.md"),
                    caption="📄 Полный одностраничник в формате Markdown"
                )
            
            # Отправляем JSON
            with open(json_path, 'rb') as f:
                await update.message.reply_document(
                    document=InputFile(f, filename=f"Данные_проекта_{session.one_pager.project_name}.json"),
                    caption="📊 Полные данные проекта в формате JSON"
                )
            
            # Генерируем и отправляем презентацию (админу)
            if update.effective_user.id == 957543705:
                try:
                    pptx_path = self.powerpoint_gen.generate(session.one_pager, session)
                    with open(pptx_path, 'rb') as f:
                        await update.message.reply_document(
                            document=InputFile(f, filename=f"Презентация_{session.one_pager.project_name}.pptx"),
                            caption="📊 Презентация проекта (админ)"
                        )
                except Exception as e:
                    logger.error(f"Ошибка генерации презентации: {e}")
                    await update.message.reply_text("⚠️ Не удалось сгенерировать презентацию.")
            
            # Финальное сообщение
            final_text = """
✅ *Сбор требований завершен!*

Что можно сделать:
• *Редактировать результат* — внести изменения в итоговые документы
• *Сформировать презентацию* — создать PowerPoint презентацию
• *Начать новый проект* — начать сбор требований для нового проекта

Файлы сохранены и доступны для скачивания.
"""
            
            await update.message.reply_text(final_text, reply_markup=get_final_keyboard())
            
        except Exception as e:
            logger.error(f"Ошибка отправки файлов: {e}")
            await update.message.reply_text(
                "⚠️ Произошла ошибка при отправке файлов. Данные сохранены на сервере.",
                reply_markup=get_final_keyboard()
            )
    
    async def handle_final_actions(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка действий в финальном состоянии"""
        text = update.message.text.strip()
        chat_id = update.effective_chat.id
        
        if text == "Редактировать результат":
            await update.message.reply_text(
                "Для редактирования используйте команды:\n"
                "/edit_name - изменить название проекта\n"
                "/edit_goal - изменить цель проекта\n"
                "/edit_all - полное редактирование\n"
                "/continue - вернуться к результатам"
            )
            return ConversationState.FINAL.value
        
        elif text == "Сформировать презентацию":
            session = self.session_manager.get_session(chat_id, update.effective_user.id)
            
            try:
                pptx_path = self.powerpoint_gen.generate(session.one_pager, session)
                
                with open(pptx_path, 'rb') as f:
                    await update.message.reply_document(
                        document=InputFile(f, filename=f"Презентация_{session.one_pager.project_name}.pptx"),
                        caption="📊 Презентация проекта"
                    )
                
                await update.message.reply_text(
                    "✅ Презентация сформирована и отправлена!",
                    reply_markup=get_final_keyboard()
                )
            except Exception as e:
                logger.error(f"Ошибка генерации презентации: {e}")
                await update.message.reply_text(
                    "⚠️ Не удалось сгенерировать презентацию. Проверьте наличие шаблона.",
                    reply_markup=get_final_keyboard()
                )
            
            return ConversationState.FINAL.value
        
        elif text == "Начать новый проект":
            return await self.fill_start(update, context)
        
        return ConversationState.FINAL.value
    
    async def handle_natural_language(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка естественного языка"""
        text = update.message.text.strip().lower()
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        # Паттерны для распознавания
        patterns = {
            r'(вернись|вернуться|назад)\s+(к|на)\s+(предыдущий|прошлый)': self._handle_back_to_previous,
            r'(вернись|вернуться|назад)\s+(к|на)\s+первостепенным': self._handle_back_to_primary,
            r'(исправь|измени|редактируй|поменяй)\s+(вопрос|ответ)\s*(\d+)': self._handle_edit_by_number,
            r'(исправь|измени|редактируй|поменяй)\s+(название|цель|задач[и]?|ситуацию|технолог|данные)': self._handle_edit_by_keyword,
            r'(хочу|надо|нужно|можно)\s+(редактировать|исправить|изменить)': self._handle_want_to_edit,
        }
        
        for pattern, handler in patterns.items():
            if re.search(pattern, text, re.IGNORECASE):
                return await handler(update, context, session, text)
            
        if session.editing_mode and session.state == ConversationState.EDITING:
            return await self.handle_editing(update, context)
        
        # Если не распознано, отправляем в обычный обработчик
        if session.state == ConversationState.PRIMARY_QUESTIONS:
            return await self.handle_primary_answer(update, context)
        elif session.state == ConversationState.CLARIFYING_QUESTIONS:
            return await self.handle_clarifying_answer(update, context)
        elif session.state == ConversationState.FINAL:
            return await self.handle_final_actions(update, context)
        
        await update.message.reply_text("Не понял команду. Пожалуйста, используйте кнопки или четкие команды.")
        return session.state.value
    
    async def _handle_back_to_previous(self, update: Update, context: ContextTypes.DEFAULT_TYPE, session: Session, text: str):
        """Обработка возврата к предыдущему вопросу"""
        if session.current_question_index > 0:
            session.current_question_index -= 1
            session.state = ConversationState.PRIMARY_QUESTIONS
            self.session_manager.save_session(session)
            
            await update.message.reply_text(
                f"Возвращаюсь к вопросу {session.current_question_index + 1}.",
                reply_markup=get_primary_keyboard()
            )
            
            return await self.ask_current_question(update, context)
        else:
            await update.message.reply_text("Это первый вопрос, назад вернуться нельзя.")
            return session.state.value
    
    async def _handle_back_to_primary(self, update: Update, context: ContextTypes.DEFAULT_TYPE, session: Session, text: str):
        """Обработка возврата к первостепенным вопросам"""
        session.state = ConversationState.PRIMARY_CONFIRMATION
        self.session_manager.save_session(session)
        
        await update.message.reply_text(
            "Возвращаюсь к первостепенным вопросам.",
            reply_markup=get_confirmation_keyboard()
        )
        
        return ConversationState.PRIMARY_CONFIRMATION.value
    
    async def _handle_edit_by_number(self, update: Update, context: ContextTypes.DEFAULT_TYPE, session: Session, text: str):
        """Редактирование по номеру вопроса"""
        numbers = re.findall(r'\d+', text)
        if numbers:
            question_num = int(numbers[0]) - 1
            
            if 0 <= question_num < len(session.primary_answers):
                session.current_question_index = question_num
                session.state = ConversationState.EDITING
                session.edit_target = "primary"
                self.session_manager.save_session(session)
                
                question_text = session.primary_answers[question_num].question
                current_answer = session.primary_answers[question_num].answer
                
                response = f"""
Редактирование вопроса {question_num + 1}:

*Вопрос:* {question_text}

*Текущий ответ:* {current_answer if current_answer else "Нет ответа"}

Пожалуйста, введите новый ответ:
"""
                await update.message.reply_text(response)
                return ConversationState.EDITING.value
        
        await update.message.reply_text("Не удалось определить номер вопроса. Попробуйте еще раз.")
        return session.state.value
    
    async def _handle_edit_by_keyword(self, update: Update, context: ContextTypes.DEFAULT_TYPE, session: Session, text: str):
        """Редактирование по ключевому слову"""
        # Поиск по ключевым словам в вопросах
        keyword_mapping = {
            'название': 0,  # Первый вопрос
            'цель': 5,      # Шестой вопрос (цель проекта)
            'задач': 8,     # Девятый вопрос (бизнес-задачи)
            'ситуацию': 3,  # Четвертый вопрос (AS IS)
            'технолог': 12, # Предпоследний вопрос (технологии)
            'данные': 6,    # Седьмой вопрос (источники данных)
        }
        
        for keyword, index in keyword_mapping.items():
            if keyword in text:
                session.current_question_index = index
                session.state = ConversationState.EDITING
                session.edit_target = "primary"
                self.session_manager.save_session(session)
                
                question_text = session.primary_answers[index].question
                current_answer = session.primary_answers[index].answer
                
                response = f"""
Нашел вопрос по ключевому слову "{keyword}":

*Вопрос {index + 1}:* {question_text}

*Текущий ответ:* {current_answer if current_answer else "Нет ответа"}

Пожалуйста, введите новый ответ:
"""
                await update.message.reply_text(response)
                return ConversationState.EDITING.value
        
        await update.message.reply_text("Не удалось найти вопрос по ключевому слову. Попробуйте указать номер.")
        return session.state.value
    
    async def _handle_want_to_edit(self, update: Update, context: ContextTypes.DEFAULT_TYPE, session: Session, text: str):
        """Обработка желания редактировать"""
        await update.message.reply_text(
            "Хотите что-то отредактировать? Укажите:\n"
            "1. Номер вопроса (1-14)\n"
            "2. Или что именно хотите изменить (название, цель, etc.)\n"
            "3. Или напишите 'редактировать все' для полного редактирования"
        )
        return session.state.value
    
    async def cancel(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Отмена текущей сессии"""
        chat_id = update.effective_chat.id
        
        if chat_id in self.session_manager.sessions:
            del self.session_manager.sessions[chat_id]
        
        # Удаляем файл сессии
        session_file = SESSIONS_DIR / f"{chat_id}.json"
        if session_file.exists():
            session_file.unlink()
        
        await update.message.reply_text(
            "Сессия отменена. Все данные удалены.\n\n"
            "Чтобы начать заново, отправьте /start",
            reply_markup=ReplyKeyboardRemove()
        )
        
        return ConversationHandler.END
    
    async def export_data(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Экспорт данных"""
        chat_id = update.effective_chat.id
        session = self.session_manager.get_session(chat_id, update.effective_user.id)
        
        if session.state != ConversationState.FINAL:
            await update.message.reply_text("Сбор требований еще не завершен. Завершите проект для экспорта.")
            return
        
        # Отправляем файлы
        json_path = SESSIONS_DIR / f"full_data_{chat_id}.json"
        md_path = SESSIONS_DIR / f"onepager_{chat_id}.md"
        
        try:
            with open(md_path, 'rb') as f:
                await update.message.reply_document(
                    document=InputFile(f, filename=f"Одностраничник_{session.one_pager.project_name}.md"),
                    caption="📄 Одностраничник проекта"
                )
            
            with open(json_path, 'rb') as f:
                await update.message.reply_document(
                    document=InputFile(f, filename=f"Данные_проекта_{session.one_pager.project_name}.json"),
                    caption="📊 Данные проекта"
                )
            
            await update.message.reply_text("✅ Экспорт завершен!")
            
        except Exception as e:
            logger.error(f"Ошибка экспорта: {e}")
            await update.message.reply_text(f"⚠️ Ошибка экспорта: {e}")
    
    async def help_command(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда помощи"""
        help_text = """
🤖 *Помощь по боту*

*Основные команды:*
/start - начать работу с ботом
/fill - начать сбор требований
/export - экспорт результатов
/cancel - отмена текущей сессии
/help - эта справка

*Естественные команды (можно писать текстом):*
- "вернись к предыдущему вопросу"
- "исправь вопрос 3"
- "измени название проекта"
- "редактируй цель"
- "хочу отредактировать"

*Кнопки:*
• Поясни - объяснение текущего вопроса
• Пропустить - пропустить вопрос
• Принять - принять предсказание бота
• Изменить - изменить предложенное предсказание
• Ввести свой - ввести свой вариант ответа

Бот анализирует ваши ответы и предлагает варианты для следующих вопросов на основе контекста.
"""
        
        await update.message.reply_text(help_text)

# ============== ЗАПУСК БОТА ==============
def main():
    """Основная функция запуска бота"""
    # Создаем экземпляр бота
    bot = RequirementsBot()
    
    # Создаем приложение
    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    # Создаем ConversationHandler
    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("start", bot.start),
            CommandHandler("fill", bot.fill_start)
        ],
        states={
            ConversationState.START.value: [
                CommandHandler("fill", bot.fill_start),
                CommandHandler("help", bot.help_command),
                CommandHandler("edit_name", bot.edit_name_command),
                CommandHandler("edit_goal", bot.edit_goal_command),
                CommandHandler("edit_all", bot.edit_all_command),
                MessageHandler(filters.TEXT & ~filters.COMMAND, bot.handle_natural_language)
            ],
            
            ConversationState.PRIMARY_QUESTIONS.value: [
                CallbackQueryHandler(bot.handle_business_process_callback, pattern=r'^bp_'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, bot.handle_primary_answer),
                CommandHandler("cancel", bot.cancel)
            ],
            
            ConversationState.WAITING_FOR_BP_DETAILS.value: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, bot.handle_primary_answer),
                CommandHandler("cancel", bot.cancel)
            ],
            
            ConversationState.PRIMARY_CONFIRMATION.value: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, bot.handle_confirmation),
                CommandHandler("cancel", bot.cancel)
            ],
            
            ConversationState.CLARIFYING_QUESTIONS.value: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, bot.handle_clarifying_answer),
                CommandHandler("cancel", bot.cancel)
            ],
            
            ConversationState.EDITING.value: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, bot.handle_edit_response),
                CommandHandler("cancel", bot.cancel)
            ],
            
            ConversationState.FINAL.value: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, bot.handle_final_actions),
                CommandHandler("export", bot.export_data),
                CommandHandler("edit_name", bot.edit_name_command),
                CommandHandler("edit_goal", bot.edit_goal_command),
                CommandHandler("edit_all", bot.edit_all_command),
                CommandHandler("continue", lambda u,c: bot._handle_back_to_primary(u,c,None,None)),
                CommandHandler("cancel", bot.cancel),
                CommandHandler("fill", bot.fill_start)
            ]
        },
        fallbacks=[
            CommandHandler("cancel", bot.cancel),
            CommandHandler("help", bot.help_command)
        ],
        allow_reentry=True,
    )
    
    # Добавляем обработчики
    application.add_handler(conv_handler)
    
    # Обработчик естественного языка (для всех сообщений, не попавших в ConversationHandler)
    application.add_handler(MessageHandler(
        filters.TEXT & ~filters.COMMAND & ~filters.Regex(r'^/'),
        bot.handle_natural_language
    ))
    
    # Команды
    application.add_handler(CommandHandler("export", bot.export_data))
    application.add_handler(CommandHandler("help", bot.help_command))
    application.add_handler(CommandHandler("cancel", bot.cancel))
    application.add_handler(CommandHandler("edit_name", bot.edit_name_command))
    application.add_handler(CommandHandler("edit_goal", bot.edit_goal_command))
    application.add_handler(CommandHandler("edit_all", bot.edit_all_command))
    application.add_handler(CommandHandler("continue", lambda u,c: bot._handle_back_to_primary(u,c,None,None)))
    
    async def continue_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
        return await bot._handle_back_to_primary(update, context)
    
    application.add_handler(CommandHandler("continue", continue_command))

    # Запуск бота
    logger.info("Бот запускается...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}", exc_info=True)
