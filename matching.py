"""Resume parsing + skill matching (pure Python, no paid APIs).

Drop resumes (.pdf .docx .tex .txt .md) in the resume/ folder. Each resume is turned into a set of skills
found in it; every job is scored against each resume and the best-matching one is reported.
"""
import re
from collections import Counter
from pathlib import Path

# canonical skill -> aliases (all lowercase). Add more via `extra_skills:` in companies.yaml.
SKILLS = {
    "python": ["python", "python3"], "java": ["java", "core java", "j2ee"], "javascript": ["javascript", "js", "es6"],
    "typescript": ["typescript"], "c++": ["c++", "cpp"], "c#": ["c#"], "c": ["c programming", "embedded c"],
    "golang": ["golang"], "rust": ["rust"], "kotlin": ["kotlin"], "swift": ["swift"], "php": ["php"], "ruby": ["ruby", "ruby on rails", "rails"],
    "scala": ["scala"], "bash": ["bash", "shell scripting", "shell script"], "sql": ["sql", "plsql", "pl/sql", "t-sql"],
    "nosql": ["nosql"], "mysql": ["mysql"], "postgresql": ["postgresql", "postgres"], "mongodb": ["mongodb", "mongo"],
    "redis": ["redis"], "sqlite": ["sqlite"], "oracle": ["oracle db", "oracle database"], "elasticsearch": ["elasticsearch", "opensearch"],
    "django": ["django", "drf", "django rest framework"], "flask": ["flask"], "fastapi": ["fastapi"], "spring boot": ["spring boot", "springboot", "spring"],
    "hibernate": ["hibernate", "jpa"], "node.js": ["node.js", "nodejs", "node"], "express": ["express.js", "expressjs", "express"],
    "nestjs": ["nestjs", "nest.js"], ".net": [".net", "dotnet", "asp.net"], "laravel": ["laravel"],
    "react": ["react", "react.js", "reactjs"], "next.js": ["next.js", "nextjs"], "angular": ["angular", "angularjs"], "vue": ["vue", "vue.js"],
    "html": ["html", "html5"], "css": ["css", "css3"], "tailwind": ["tailwind", "tailwindcss"],
    "rest api": ["rest", "rest api", "rest apis", "restful", "restful api", "restful apis", "api development"], "graphql": ["graphql"], "grpc": ["grpc"],
    "microservices": ["microservices", "microservice"], "websockets": ["websocket", "websockets"], "jwt": ["jwt"], "oauth": ["oauth", "oauth2"],
    "rbac": ["rbac"], "celery": ["celery"], "rabbitmq": ["rabbitmq"], "kafka": ["kafka"], "multithreading": ["multithreading", "concurrency"],
    "docker": ["docker", "containerization"], "kubernetes": ["kubernetes", "k8s"], "aws": ["aws", "amazon web services", "ec2", "s3", "lambda"],
    "azure": ["azure"], "gcp": ["gcp", "google cloud"], "linux": ["linux", "unix"], "git": ["git", "github", "gitlab", "version control"],
    "ci/cd": ["ci/cd", "cicd", "jenkins", "github actions", "gitlab ci"], "terraform": ["terraform"], "ansible": ["ansible"], "devops": ["devops"],
    "machine learning": ["machine learning", "ml"], "deep learning": ["deep learning"], "nlp": ["nlp", "natural language processing"],
    "llm": ["llm", "llms", "large language model", "large language models", "generative ai", "genai", "gen ai"], "rag": ["rag", "retrieval augmented generation", "retrieval-augmented generation"],
    "langchain": ["langchain", "langgraph", "llamaindex"], "agentic ai": ["agentic ai", "ai agents", "ai agent", "agentic"], "openai": ["openai", "gemini api"],
    "ai": ["ai", "artificial intelligence"], "pytorch": ["pytorch"], "tensorflow": ["tensorflow", "keras"], "scikit-learn": ["scikit-learn", "sklearn"],
    "pandas": ["pandas"], "numpy": ["numpy"], "opencv": ["opencv"], "computer vision": ["computer vision"],
    "data structures": ["data structures", "dsa"], "algorithms": ["algorithms", "algorithm"], "oop": ["oop", "oops", "object oriented", "object-oriented"],
    "system design": ["system design", "low level design", "high level design"], "design patterns": ["design patterns"],
    "unit testing": ["unit testing", "unit tests", "tdd", "pytest", "junit", "test driven development"], "selenium": ["selenium"], "postman": ["postman"],
    "agile": ["agile", "scrum"], "jira": ["jira"], "sdlc": ["sdlc"],
    "power bi": ["power bi", "powerbi"], "tableau": ["tableau"], "excel": ["excel", "ms excel", "advanced excel"],
    "data analysis": ["data analysis", "data analytics", "data analyst"], "etl": ["etl", "data pipelines", "data pipeline"], "data warehouse": ["data warehouse", "data warehousing"],
    "spark": ["spark", "pyspark"], "hadoop": ["hadoop"], "airflow": ["airflow"], "snowflake": ["snowflake"], "dbt": ["dbt"],
    "cobol": ["cobol"], "jcl": ["jcl"], "db2": ["db2"], "mainframe": ["mainframe", "mainframes", "z/os", "cics"],
    "cybersecurity": ["cybersecurity", "cyber security", "infosec", "information security", "siem", "penetration testing", "owasp", "vulnerability"],
    "blockchain": ["blockchain", "web3"], "solidity": ["solidity", "smart contracts"], "iot": ["iot", "internet of things"],
    "embedded": ["embedded", "firmware", "rtos"], "android": ["android"], "ios": ["ios"], "flutter": ["flutter"],
    "networking": ["tcp/ip", "networking", "http", "dns"], "api": ["api", "apis"],
}


def compile_skills(extra=()):
    vocab = {k: list(v) for k, v in SKILLS.items()}
    for e in extra or []:
        vocab.setdefault(e.lower(), []).append(e.lower())
    out = []
    for canon, aliases in vocab.items():
        alt = "|".join(re.escape(a) for a in sorted(set(aliases + [canon]), key=len, reverse=True))
        out.append((canon, re.compile(rf"(?<![a-z0-9+#.]){'(?:' + alt + ')'}(?![a-z0-9+#])")))
    return out


def extract(text, compiled):
    t = text.lower()
    c = Counter()
    for canon, rx in compiled:
        n = len(rx.findall(t))
        if n:
            c[canon] = n
    return c


# ---------------- resume loading ----------------
def _read(path):
    suf = path.suffix.lower()
    if suf == ".pdf":
        from pypdf import PdfReader
        return " ".join((p.extract_text() or "") for p in PdfReader(str(path)).pages)
    if suf == ".docx":
        import docx
        d = docx.Document(str(path))
        parts = [p.text for p in d.paragraphs]
        for tb in d.tables:
            parts += [c.text for r in tb.rows for c in r.cells]
        return "\n".join(parts)
    txt = path.read_text(errors="ignore")
    if suf == ".tex":
        txt = re.sub(r"(?<!\\)%.*", " ", txt)
        txt = re.sub(r"\\[a-zA-Z]+\*?(\[[^\]]*\])?", " ", txt)
        txt = re.sub(r"[{}\\]", " ", txt)
    return txt


def load_resumes(folder, compiled, log=print):
    """-> {resume file name: set of skills}"""
    out = {}
    for p in sorted(Path(folder).glob("*")):
        if p.suffix.lower() not in (".pdf", ".docx", ".tex", ".txt", ".md") or p.name.lower().startswith("readme"):
            continue
        try:
            text = _read(p).replace("ﬁ", "fi").replace("ﬂ", "fl")
            skills = set(extract(text, compiled))
            out[p.name] = skills
            log(f"resume {p.name}: {len(skills)} skills -> {', '.join(sorted(skills)[:12])}{' ...' if len(skills) > 12 else ''}")
        except Exception as e:
            log(f"could not read resume {p.name}: {e}")
    return out


# ---------------- scoring ----------------
def score_job(text, resumes, compiled):
    """Score a job (title + JD text) against every resume; return the best one.
    -> dict(score 0-100, best resume, matched skills, missing skills)"""
    jd = extract(text, compiled)
    best = dict(score=0, best="", matched=[], missing=[])
    if not jd:
        return best
    for name, rs in resumes.items():
        matched = [s for s, _ in jd.most_common() if s in rs]
        cov = len(matched) / len(jd)
        sc = round(100 * (0.6 * cov + 0.4 * min(len(matched), 6) / 6))
        if sc > best["score"] or not best["best"]:
            best = dict(score=sc, best=name, matched=matched,
                        missing=[s for s, _ in jd.most_common() if s not in rs][:10])
    return best
