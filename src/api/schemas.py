# -*- coding: utf-8 -*-
"""
Schemas Pydantic v2 para validação estrita de payloads de entrada na API REST.
"""

import re
from typing import List, Optional
from pydantic import BaseModel, Field, field_validator


class ChamadoUpdatePayload(BaseModel):
    """Payload de atualização parcial (PATCH) ou completa (PUT) de um chamado."""

    status: Optional[str] = Field(default=None, max_length=50)
    titulo: Optional[str] = Field(default=None, max_length=500)
    hostname: Optional[str] = Field(default=None, max_length=255)
    ip_origem: Optional[str] = Field(default=None, max_length=100)
    tag: Optional[str] = Field(default=None, max_length=120)
    cidade_predio: Optional[str] = Field(default=None, max_length=255)
    unidade: Optional[str] = Field(default=None, max_length=255)
    localidade_fisica: Optional[str] = Field(default=None, max_length=255)
    descricao: Optional[str] = Field(default=None, max_length=25000)
    andamento: Optional[str] = Field(default=None, max_length=25000)
    tag_manual: Optional[bool] = Field(default=None)
    dados_manuais: Optional[bool] = Field(default=None)

    @field_validator("status")
    @classmethod
    def validate_status(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return None
        cleaned = v.strip()
        if not cleaned:
            raise ValueError("O campo 'status' não pode ser vazio quando informado.")
        return cleaned

    @field_validator("hostname", "ip_origem")
    @classmethod
    def sanitize_network_fields(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return None
        cleaned = v.strip()
        if cleaned and not re.fullmatch(r"[a-zA-Z0-9.\-_:/\s]{1,255}", cleaned):
            raise ValueError("Valor contém caracteres não permitidos em campo de rede.")
        return cleaned


class ComentarioCreatePayload(BaseModel):
    """Payload para registrar novo comentário/nota técnica em um chamado."""

    autor: str = Field(..., min_length=1, max_length=150)
    texto: str = Field(..., min_length=1, max_length=10000)
    data: Optional[str] = Field(default=None, max_length=40)

    @field_validator("autor", "texto")
    @classmethod
    def must_not_be_blank(cls, v: str) -> str:
        cleaned = v.strip()
        if not cleaned:
            raise ValueError("Campo obrigatório não pode conter apenas espaços.")
        return cleaned


class ComentariosBatchPayload(BaseModel):
    """Payload para consulta de comentários em lote por lista de IDs."""

    chamado_ids: List[str] = Field(..., min_length=1, max_length=500)


class CronUpdatePayload(BaseModel):
    """Payload para ativar/desativar ou ajustar parâmetros de uma rotina do Cron."""

    ativo: Optional[bool] = None
    tipo_agendamento: Optional[str] = Field(default=None, max_length=30)
    intervalo_valor: Optional[int] = Field(default=None, ge=1, le=365)
    intervalo_unidade: Optional[str] = Field(default=None, max_length=20)
    horario_fixo: Optional[str] = Field(default=None, max_length=10)
    apenas_dias_uteis: Optional[bool] = None

    @field_validator("tipo_agendamento")
    @classmethod
    def validate_tipo(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return None
        cleaned = v.strip().lower()
        if cleaned not in {"intervalo", "horario_fixo"}:
            raise ValueError("tipo_agendamento deve ser 'intervalo' ou 'horario_fixo'.")
        return cleaned

    @field_validator("intervalo_unidade")
    @classmethod
    def validate_unidade(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return None
        cleaned = v.strip().lower()
        if cleaned not in {"minutos", "horas", "dias"}:
            raise ValueError("intervalo_unidade deve ser 'minutos', 'horas' ou 'dias'.")
        return cleaned

    @field_validator("horario_fixo")
    @classmethod
    def validate_horario(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return None
        cleaned = v.strip()
        if not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", cleaned):
            raise ValueError("horario_fixo deve estar no formato HH:MM (ex: 08:30).")
        return cleaned
