
"use client";

import React, { useState, useCallback, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import iconv from 'iconv-lite'; // Import iconv-lite for encoding
import { parse, format, isValid, subMonths, isValid as dateFnsIsValid } from 'date-fns'; // Import date-fns functions

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectGroup, SelectItem, SelectLabel, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { Upload, Settings, ArrowRight, Trash2, Plus, HelpCircle, Columns, Edit, Code, Loader2, Save, RotateCcw, ArrowUp, ArrowDown, Calculator, Server, Info, Download, LayoutList, AlertTriangle } from 'lucide-react'; // Added Info, Download, LayoutList, AlertTriangle icons
import { useToast } from "@/hooks/use-toast";
import { Textarea } from '@/components/ui/textarea';
import { Switch } from "@/components/ui/switch";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription, DialogFooter, DialogClose, DialogTrigger } from "@/components/ui/dialog"; // Import Dialog components
import { Checkbox } from '@/components/ui/checkbox'; // Import Checkbox
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover"; // Import Popover


// Define types
type DataType = 'Inteiro' | 'Alfanumérico' | 'Numérico' | 'Data' | 'CPF' | 'CNPJ';
type PredefinedField = {
    id: string;
    name: string;
    isCore: boolean; // True for original, hardcoded fields
    comment?: string;
    isPersistent?: boolean; // Tracks if a custom field is saved in localStorage
    group?: string; // For grouping in select
};
type ColumnMapping = {
  originalHeader: string;
  mappedField: string | null; // ID of predefined field or null
  dataType: DataType | null;
  length?: number | null;
  removeMask: boolean;
};
type OutputFormat = 'txt' | 'csv';
type PaddingDirection = 'left' | 'right';
type DateFormat = 'YYYYMMDD' | 'DDMMYYYY';
type OutputEncoding = 'UTF-8' | 'ISO-8859-1' | 'Windows-1252';

// Calculated Field Type
type CalculatedFieldType = 'CalculateStartDate' | 'CalculateSituacaoRetorno' | 'FormatPeriodMMAAAA';
type CalculatedFieldConfig = {
    id: string; // Unique ID
    type: CalculatedFieldType;
    fieldName: string; // Name for display
    order: number;
    requiredInputFields: string[]; // Mapped field IDs required for calculation
    parameters?: Record<string, any>; // e.g., { period: 'DD/MM/AAAA' }
    dateFormat?: DateFormat; // For Data type outputs
    length?: number; // Required for TXT
    paddingChar?: string; // For TXT
    paddingDirection?: PaddingDirection; // For TXT
};

// Consolidated Output Field Type using discriminated union
type OutputFieldConfig = {
  id: string; // Unique ID for React key prop
  order: number;
  length?: number; // Required for TXT
  paddingChar?: string; // For TXT
  paddingDirection?: PaddingDirection; // For TXT
  dateFormat?: DateFormat; // For Data type fields
  isStaticFromPreset?: boolean; // Flag to identify static fields originating from a preset
} & (
  | { isStatic: false; isCalculated: false; mappedField: string } // Mapped field
  | { isStatic: true; isCalculated: false; fieldName: string; staticValue: string } // Static field
  | { isStatic: false; isCalculated: true; type: CalculatedFieldType; fieldName: string; requiredInputFields: string[]; parameters?: Record<string, any> } // Calculated field
);


type OutputConfig = {
  name?: string; // Optional name for saved configuration
  format: OutputFormat;
  delimiter?: string; // For CSV
  fields: OutputFieldConfig[];
  encoding: OutputEncoding; // Added encoding to the config
};


// Static Field Dialog State
type StaticFieldDialogState = {
    isOpen: boolean;
    isEditing: boolean;
    fieldId?: string; // ID of the field being edited
    fieldName: string;
    staticValue: string;
    length: string; // Use string for input control
    paddingChar: string;
    paddingDirection: PaddingDirection;
}

// Predefined Field Dialog State
type PredefinedFieldDialogState = {
    isOpen: boolean;
    isEditing: boolean;
    fieldId?: string;
    fieldName: string;
    isPersistent: boolean; // Replaced 'persist' for clarity
    comment: string;
}

// Calculated Field Dialog State
type CalculatedFieldDialogState = {
    isOpen: boolean;
    isEditing: boolean;
    fieldId?: string;
    fieldName: string;
    type: CalculatedFieldType | '';
    parameters: {
        period?: string; // For CalculateStartDate
    };
    requiredInputFields: {
        parcelasPagas?: string | null; // Mapped field ID for 'PARCELAS_PAGAS' for CalculateStartDate
        valorRealizado?: string | null; // Mapped field ID for 'VALOR_REALIZADO' for CalculateSituacaoRetorno
    };
    length: string;
    paddingChar: string;
    paddingDirection: PaddingDirection;
    dateFormat: DateFormat | '';
}

// Save/Load Configuration Dialog State
type ConfigManagementDialogState = {
    isOpen: boolean;
    action: 'save' | 'load' | null;
    configName: string;
    selectedConfigToLoad: string | null;
}


const CORE_PREDEFINED_FIELDS_UNSORTED: PredefinedField[] = [
  { id: 'matricula', name: 'Matrícula', isCore: true, comment: 'Número de matrícula do servidor/funcionário.', group: 'Padrão', isPersistent: true },
  { id: 'cpf', name: 'CPF', isCore: true, comment: 'Cadastro de Pessoa Física. Será formatado sem máscara na saída se a opção estiver marcada.', group: 'Padrão', isPersistent: true },
  { id: 'rg', name: 'RG', isCore: true, comment: 'Registro Geral (Identidade). Pode conter letras e números.', group: 'Padrão', isPersistent: true },
  { id: 'nome', name: 'Nome', isCore: true, comment: 'Nome completo.', group: 'Padrão', isPersistent: true },
  { id: 'email', name: 'E-mail', isCore: true, comment: 'Endereço de e-mail.', group: 'Margem', isPersistent: true },
  { id: 'cnpj', name: 'CNPJ', isCore: true, comment: 'Cadastro Nacional da Pessoa Jurídica. Será formatado sem máscara na saída se a opção estiver marcada.', group: 'Padrão', isPersistent: true },
  { id: 'regime', name: 'Regime', isCore: true, comment: 'Regime de contratação (ex: CLT, Estatutário).', group: 'Margem', isPersistent: true },
  { id: 'situacao_usuario', name: 'Situação do Usuário', isCore: true, comment: 'Situação do usuário/servidor (ex: Ativo, Inativo, Licença).', group: 'Margem', isPersistent: true },
  { id: 'categoria', name: 'Categoria', isCore: true, comment: 'Categoria funcional.', group: 'Margem', isPersistent: true },
  { id: 'secretaria', name: 'Local de Trabalho', isCore: true, comment: 'Secretaria ou órgão de lotação.', group: 'Margem', isPersistent: true }, // Renamed from Secretaria
  { id: 'setor', name: 'Setor', isCore: true, comment: 'Setor ou departamento específico (usado para Lotação em alguns layouts).', group: 'Margem', isPersistent: true },
  { id: 'margem_bruta', name: 'Margem Bruta', isCore: true, comment: 'Valor da margem bruta consignável (Numérico).', group: 'Margem', isPersistent: true },
  { id: 'margem_reservada', name: 'Margem Reservada', isCore: true, comment: 'Valor da margem reservada (Numérico). Usado como Margem2 em layouts de cartão.', group: 'Margem', isPersistent: true },
  { id: 'margem_liquida', name: 'Margem Líquida', isCore: true, comment: 'Valor da margem líquida disponível (Numérico). Usado como Margem ou Margem1.', group: 'Margem', isPersistent: true },
  { id: 'parcelas_pagas', name: 'Parcelas Pagas', isCore: true, comment: 'Número de parcelas pagas de um contrato/empréstimo (Inteiro).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'data_nascimento', name: 'Data de Nascimento', isCore: true, comment: 'Data de nascimento do indivíduo (Data).', group: 'Margem', isPersistent: true },
  { id: 'data_admissao', name: 'Data de Admissão', isCore: true, comment: 'Data de admissão na empresa/órgão (Data).', group: 'Margem', isPersistent: true },
  { id: 'data_fim_contrato', name: 'Data Fim do Contrato', isCore: true, comment: 'Data de término do contrato, se aplicável (Data).', group: 'Margem', isPersistent: true },
  { id: 'sinal_margem', name: 'Sinal da Margem', isCore: true, comment: 'Sinal indicativo da margem (+ ou -).', group: 'Margem', isPersistent: true },
  { id: 'estabelecimento_empresa', name: 'Estabelecimento/Empresa', isCore: true, comment: 'Nome do estabelecimento ou empresa que concede o crédito/serviço.', group: 'Padrão', isPersistent: true }, // Corrected name
  { id: 'consignataria_contrato', name: 'Consignatária (Contrato)', isCore: true, comment: 'Nome da empresa consignatária específica do contrato (usado em layouts de Histórico/Retorno).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'orgao_filial', name: 'Órgão/Filial', isCore: true, comment: 'Nome do órgão ou filial.', group: 'Padrão', isPersistent: true },
  { id: 'verba_rubrica', name: 'Verba/Rubrica', isCore: true, comment: 'Código da verba ou rubrica.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'prazo_total', name: 'Prazo Total', isCore: true, comment: 'Prazo total de um contrato/empréstimo em meses (Inteiro).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'parcelas_restantes', name: 'Parcelas Restantes', isCore: true, comment: 'Número de parcelas restantes de um contrato/empréstimo (Inteiro).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'valor_parcela', name: 'Valor da Parcela', isCore: true, comment: 'Valor de cada parcela (Numérico). Usado como "Valor" no layout Histórico.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'valor_financiado', name: 'Valor Financiado', isCore: true, comment: 'Valor total financiado (Numérico).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'valor_total', name: 'Valor Total', isCore: true, comment: 'Valor total de uma transação ou contrato (Numérico).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'cet_mensal', name: 'CET Mensal', isCore: true, comment: 'Custo Efetivo Total Mensal (Numérico).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'cet_anual', name: 'CET Anual', isCore: true, comment: 'Custo Efetivo Total Anual (Numérico).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'numero_contrato', name: 'Número do Contrato', isCore: true, comment: 'Número identificador do contrato.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'verba_rubrica_ferias', name: 'Verba/Rubrica Férias', isCore: true, comment: 'Código da verba/rubrica de férias.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'valor_previsto', name: 'Valor Previsto', isCore: true, comment: 'Valor previsto de um lançamento (Numérico).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'valor_realizado', name: 'Valor Realizado', isCore: true, comment: 'Valor realizado de um lançamento (Numérico).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'observacao', name: 'Observação', isCore: true, comment: 'Observações gerais. Usado como "Motivo" no layout Retorno.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'motivo_critica', name: 'Motivo(Crítica)', isCore: true, comment: 'Motivo da crítica, erro ou observação detalhada para um registro ou parcela.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'situacao_parcela', name: 'Situação Parcela', isCore: true, comment: 'Situação de uma parcela (ex: Paga, Aberta, Vencida).', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'periodo', name: 'Período', isCore: true, comment: 'Período de referência (Data). Usado como base para "Período MMAAAA" no layout Retorno.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'identificador', name: 'Identificador', isCore: true, comment: 'Identificador único genérico.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'indice', name: 'Índice', isCore: true, comment: 'Valor de índice ou fator.', group: 'Histórico/Retorno', isPersistent: true },
  { id: 'tempo_casa', name: 'Tempo de Casa', isCore: true, comment: 'Tempo de serviço na empresa/órgão.', group: 'Margem', isPersistent: true },
  { id: 'situacao', name: 'Situação', isCore: true, comment: 'Situação funcional (ex: Ativo, Inativo). Esta será substituída por Situação do Usuário.', group: 'Margem', isPersistent: true },
].map(f => ({ ...f, isPersistent: true }));

const CORE_PREDEFINED_FIELDS = CORE_PREDEFINED_FIELDS_UNSORTED.sort((a, b) => a.name.localeCompare(b.name));


const DATA_TYPES: DataType[] = ['Inteiro', 'Alfanumérico', 'Numérico', 'Data', 'CPF', 'CNPJ'];
const OUTPUT_ENCODINGS: OutputEncoding[] = ['UTF-8', 'ISO-8859-1', 'Windows-1252'];
const DATE_FORMATS: DateFormat[] = ['YYYYMMDD', 'DDMMYYYY'];

const NONE_VALUE_PLACEHOLDER = "__NONE__";
const PREDEFINED_FIELDS_STORAGE_KEY = 'sca-predefined-fields-v1.2';
const SAVED_CONFIGS_STORAGE_KEY = 'sca-saved-configs-v1.2';

// Layout Preset Types and Definitions
type PresetOutputField = {
    targetMappedId?: string;
    staticFieldName?: string;
    staticValue?: string;
    calculatedInternalName: string;
    calculatedDisplayName?: string;
    calculationType?: CalculatedFieldType;
    requiredInputFieldsForCalc?: string[];
    length: number;
    paddingChar: string;
    paddingDirection: PaddingDirection;
    dateFormat?: DateFormat;
};

type LayoutPresetDefinition = {
    id: string;
    name: string;
    fields: PresetOutputField[];
};

const LAYOUT_PRESETS: LayoutPresetDefinition[] = [
    {
        id: "margem_simples_econsig",
        name: "Layout para Margem simples (Padrão eConsig)",
        fields: [
            { calculatedInternalName: "Matricula", targetMappedId: 'matricula', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "CPF", targetMappedId: 'cpf', length: 11, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Nome", targetMappedId: 'nome', length: 50, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "Estabelecimento", targetMappedId: 'estabelecimento_empresa', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Orgao", targetMappedId: 'orgao_filial', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Margem", targetMappedId: 'margem_liquida', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "DataNascimento", targetMappedId: 'data_nascimento', length: 8, paddingChar: '0', paddingDirection: 'left', dateFormat: 'DDMMYYYY' },
            { calculatedInternalName: "DataAdmissao", targetMappedId: 'data_admissao', length: 8, paddingChar: '0', paddingDirection: 'left', dateFormat: 'DDMMYYYY' },
            { calculatedInternalName: "DataFimContrato", targetMappedId: 'data_fim_contrato', length: 8, paddingChar: '0', paddingDirection: 'left', dateFormat: 'DDMMYYYY' },
            { calculatedInternalName: "Regime", targetMappedId: 'regime', length: 40, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "Lotacao", targetMappedId: 'secretaria', length: 40, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "RG", targetMappedId: 'rg', length: 15, paddingChar: ' ', paddingDirection: 'left' },
            { calculatedInternalName: "Email", targetMappedId: 'email', length: 100, paddingChar: ' ', paddingDirection: 'right' },
        ]
    },
    {
        id: "margem_cartao_econsig",
        name: "Layout para Margem simples + cartão (Padrão eConsig)",
        fields: [
            { calculatedInternalName: "Matricula", targetMappedId: 'matricula', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "CPF", targetMappedId: 'cpf', length: 11, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Nome", targetMappedId: 'nome', length: 50, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "Estabelecimento", targetMappedId: 'estabelecimento_empresa', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Orgao", targetMappedId: 'orgao_filial', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Margem1", targetMappedId: 'margem_liquida', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Margem2", targetMappedId: 'margem_reservada', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "DataNascimento", targetMappedId: 'data_nascimento', length: 8, paddingChar: '0', paddingDirection: 'left', dateFormat: 'DDMMYYYY' },
            { calculatedInternalName: "DataAdmissao", targetMappedId: 'data_admissao', length: 8, paddingChar: '0', paddingDirection: 'left', dateFormat: 'DDMMYYYY' },
            { calculatedInternalName: "DataFimContrato", targetMappedId: 'data_fim_contrato', length: 8, paddingChar: '0', paddingDirection: 'left', dateFormat: 'DDMMYYYY' },
            { calculatedInternalName: "Regime", targetMappedId: 'regime', length: 40, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "Lotacao", targetMappedId: 'secretaria', length: 40, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "RG", targetMappedId: 'rg', length: 15, paddingChar: ' ', paddingDirection: 'left' },
            { calculatedInternalName: "Email", targetMappedId: 'email', length: 100, paddingChar: ' ', paddingDirection: 'right' },
        ]
    },
    {
        id: "historico_econsig",
        name: "Layout para Histórico (Padrão eConsig)",
        fields: [
            { calculatedInternalName: "Matricula", targetMappedId: 'matricula', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "CPF", targetMappedId: 'cpf', length: 11, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Nome", targetMappedId: 'nome', length: 50, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "Estabelecimento", targetMappedId: 'estabelecimento_empresa', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Orgao", targetMappedId: 'orgao_filial', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Verba", targetMappedId: 'verba_rubrica', length: 5, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Prazo", targetMappedId: 'prazo_total', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "ParcelasPagas", targetMappedId: 'parcelas_pagas', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Valor", targetMappedId: 'valor_parcela', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "DataInicioContrato", calculatedDisplayName: "Data Início Contrato (Calc.)", calculationType: 'CalculateStartDate', requiredInputFieldsForCalc: ['parcelas_pagas'], length: 8, paddingChar: '0', paddingDirection: 'left', dateFormat: 'DDMMYYYY' },
            { calculatedInternalName: "Contrato", targetMappedId: 'numero_contrato', length: 30, paddingChar: ' ', paddingDirection: 'right' },
        ]
    },
    {
        id: "retorno_simples_econsig",
        name: "Layout para Retorno simples (Padrão eConsig)",
        fields: [
            { calculatedInternalName: "Matricula", targetMappedId: 'matricula', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "CPF", targetMappedId: 'cpf', length: 11, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Nome", targetMappedId: 'nome', length: 50, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "Estabelecimento", targetMappedId: 'estabelecimento_empresa', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Orgao", targetMappedId: 'orgao_filial', length: 3, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Verba", targetMappedId: 'verba_rubrica', length: 5, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "ValorPrevisto", targetMappedId: 'valor_previsto', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "ValorRealizado", targetMappedId: 'valor_realizado', length: 10, paddingChar: '0', paddingDirection: 'left' },
            { calculatedInternalName: "Motivo", targetMappedId: 'observacao', length: 100, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "SituacaoRetorno", calculatedDisplayName: "Situação (Calc.)", calculationType: 'CalculateSituacaoRetorno', requiredInputFieldsForCalc: ['valor_realizado'], length: 1, paddingChar: ' ', paddingDirection: 'right' },
            { calculatedInternalName: "PeriodoMMAAAA", calculatedDisplayName: "Período MMAAAA (Calc.)", calculationType: 'FormatPeriodMMAAAA', requiredInputFieldsForCalc: ['periodo'], length: 6, paddingChar: '0', paddingDirection: 'left', dateFormat: 'DDMMYYYY' }, // Note: dateFormat here is for internal consistency, output is MMyyyy
        ]
    }
];


// Helper to check if a data type is numeric-like
const isNumericType = (dataType: DataType | null): boolean => {
    return dataType === 'Inteiro' || dataType === 'Numérico' || dataType === 'CPF' || dataType === 'CNPJ';
}

// Helper to get default padding char based on type
const getDefaultPaddingChar = (field: OutputFieldConfig, mappings: ColumnMapping[]): string => {
    if (field.isStatic) {
        // Default to space for static unless value is purely numeric
        return /^-?\d+$/.test(field.staticValue) ? '0' : ' ';
    } else if (field.isCalculated) {
         // Calculated fields default based on expected output
         if (field.type === 'FormatPeriodMMAAAA' || field.type === 'CalculateStartDate') return '0'; // Date/period as number like
         if (field.type === 'CalculateSituacaoRetorno') return ' '; // T or R
         return ' ';
    } else {
        const mapping = mappings.find(m => m.mappedField === field.mappedField);
        return isNumericType(mapping?.dataType ?? null) ? '0' : ' ';
    }
}

// Helper to get default padding direction based on type
const getDefaultPaddingDirection = (field: OutputFieldConfig, mappings: ColumnMapping[]): PaddingDirection => {
     if (field.isStatic) {
        // Default to left for static if value is purely numeric
        return /^-?\d+$/.test(field.staticValue) ? 'left' : 'right';
    } else if (field.isCalculated) {
         // Calculated fields: Left for numeric-like outputs
         if (field.type === 'FormatPeriodMMAAAA' || field.type === 'CalculateStartDate') return 'left';
         if (field.type === 'CalculateSituacaoRetorno') return 'right'; // T or R
         return 'right';
    } else {
        const mapping = mappings.find(m => m.mappedField === field.mappedField);
        return isNumericType(mapping?.dataType ?? null) ? 'left' : 'right';
    }
}

// Helper to format number to specified decimals (handles negatives)
const formatNumber = (value: string | number, decimals: number): string => {
    const cleanedValue = String(value).replace(/[R$., ]/g, (match) => match === '.' ? '.' : '');
    const num = Number(cleanedValue.replace(',', '.'));

    if (isNaN(num)) return '';
    return num.toFixed(decimals);
}

// Helper to remove mask based on type
const removeMaskHelper = (value: string, dataType: DataType | null): string => {
    if (!dataType || value === null || value === undefined) return '';
    const stringValue = String(value);

    switch (dataType) {
        case 'CPF':
        case 'CNPJ':
        case 'Inteiro':
        case 'Numérico':
            return stringValue.replace(/\D/g, ''); // Remove all non-digits
        case 'RG':
            return stringValue.replace(/[.-]/g, '');
        case 'Data':
            return stringValue.replace(/\D/g, '');
        case 'Alfanumérico':
        default:
            return stringValue;
    }
}

// Download Dialog State
type DownloadDialogState = {
    isOpen: boolean;
    proposedFilename: string;
    finalFilename: string;
}

export default function Home() {
  const { toast } = useToast();
  const [file, setFile] = useState<File | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileData, setFileData] = useState<any[]>([]);
  const [columnMappings, setColumnMappings] = useState<ColumnMapping[]>([]);
  const [outputConfig, setOutputConfig] = useState<OutputConfig>({ name: 'Configuração Atual', format: 'txt', fields: [], encoding: 'UTF-8' });
  const [predefinedFields, setPredefinedFields] = useState<PredefinedField[]>([]); // Initialized in useEffect
  const [newFieldName, setNewFieldName] = useState<string>('');
  const [convertedData, setConvertedData] = useState<string | Buffer>('');
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingMessage, setProcessingMessage] = useState<string>('Processando...');
  const [activeTab, setActiveTab] = useState<string>("upload");
  const [showPreview, setShowPreview] = useState<boolean>(false);
  const [selectedLayoutPresetId, setSelectedLayoutPresetId] = useState<string>("custom");

  const [staticFieldDialogState, setStaticFieldDialogState] = useState<StaticFieldDialogState>({
        isOpen: false,
        isEditing: false,
        fieldName: '',
        staticValue: '',
        length: '',
        paddingChar: ' ',
        paddingDirection: 'right',
    });
   const [predefinedFieldDialogState, setPredefinedFieldDialogState] = useState<PredefinedFieldDialogState>({
       isOpen: false,
       isEditing: false,
       fieldName: '',
       isPersistent: false, // Initialize as not persistent by default for new fields
       comment: '',
   });
    const [downloadDialogState, setDownloadDialogState] = useState<DownloadDialogState>({
        isOpen: false,
        proposedFilename: '',
        finalFilename: '',
    });
    const [calculatedFieldDialogState, setCalculatedFieldDialogState] = useState<CalculatedFieldDialogState>({
        isOpen: false,
        isEditing: false,
        fieldName: '',
        type: '',
        parameters: {},
        requiredInputFields: { parcelasPagas: null, valorRealizado: null },
        length: '',
        paddingChar: ' ',
        paddingDirection: 'right',
        dateFormat: '',
    });
    const [configManagementDialogState, setConfigManagementDialogState] = useState<ConfigManagementDialogState>({
        isOpen: false,
        action: null,
        configName: '',
        selectedConfigToLoad: null,
    });
    const [savedConfigs, setSavedConfigs] = useState<OutputConfig[]>([]); // State for saved configurations


  const appVersion = process.env.NEXT_PUBLIC_APP_VERSION || '1.2.5'; // Set current version

   // Load predefined fields and saved configurations from localStorage on mount
   useEffect(() => {
       // Load Predefined Fields
       const storedFieldsJson = localStorage.getItem(PREDEFINED_FIELDS_STORAGE_KEY);
       let customFields: PredefinedField[] = [];
       if (storedFieldsJson) {
           try {
               customFields = JSON.parse(storedFieldsJson)
                    .filter((f: any) => typeof f === 'object' && f.id && f.name && !f.isCore) // Filter out core fields just in case
                    .map((f: any) => ({ // Map to ensure correct structure and add isPersistent flag
                        id: f.id,
                        name: f.name,
                        isCore: false,
                        comment: f.comment || '',
                        isPersistent: true, // Fields from storage are persistent
                        group: f.group || 'Personalizado' // Assign a default group if not present
                    }));
           } catch (e) {
               console.error("Falha ao analisar campos pré-definidos do localStorage:", e);
               localStorage.removeItem(PREDEFINED_FIELDS_STORAGE_KEY);
           }
       }
       const combined = [...CORE_PREDEFINED_FIELDS];
       const coreIds = new Set(CORE_PREDEFINED_FIELDS.map(f => f.id));
       customFields.forEach(cf => {
           if (!coreIds.has(cf.id) && !combined.some(f => f.id === cf.id)) {
               combined.push(cf);
           }
       });
       setPredefinedFields(combined.sort((a, b) => a.name.localeCompare(b.name))); // Sort all fields

        // Load Saved Configurations
        const storedConfigsJson = localStorage.getItem(SAVED_CONFIGS_STORAGE_KEY);
        if (storedConfigsJson) {
            try {
                const loadedConfigs = JSON.parse(storedConfigsJson) as OutputConfig[];
                // Basic validation: ensure it's an array and items have a name and fields
                if (Array.isArray(loadedConfigs) && loadedConfigs.every(c => c.name && Array.isArray(c.fields))) {
                    setSavedConfigs(loadedConfigs);
                } else {
                    console.warn("Formato inválido encontrado no localStorage para configurações salvas. Ignorando.");
                    localStorage.removeItem(SAVED_CONFIGS_STORAGE_KEY);
                }
            } catch (e) {
                console.error("Falha ao analisar configurações salvas do localStorage:", e);
                localStorage.removeItem(SAVED_CONFIGS_STORAGE_KEY);
            }
        }

   }, []);

   // Save only custom, persistent predefined fields to localStorage
   const saveCustomPredefinedFields = useCallback((fieldsToSave: PredefinedField[]) => {
       // Filter for non-core fields marked as persistent
       const customPersistentFields = fieldsToSave.filter(f => !f.isCore && f.isPersistent);
       try {
           localStorage.setItem(PREDEFINED_FIELDS_STORAGE_KEY, JSON.stringify(customPersistentFields));
       } catch (e) {
           console.error("Falha ao salvar campos pré-definidos no localStorage:", e);
           toast({ title: "Erro", description: "Falha ao salvar campos pré-definidos personalizados.", variant: "destructive" });
       }
   }, [toast]);

   // Save configurations to localStorage
   const saveAllConfigs = useCallback((configsToSave: OutputConfig[]) => {
       try {
           localStorage.setItem(SAVED_CONFIGS_STORAGE_KEY, JSON.stringify(configsToSave));
       } catch (e) {
           console.error("Falha ao salvar configurações no localStorage:", e);
           toast({ title: "Erro", description: "Falha ao salvar configurações.", variant: "destructive" });
       }
   }, [toast]);

   // Function to get sample data for preview
   const getSampleData = (): any[] => {
       return fileData.slice(0, 5);
   };

   const resetSelectedLayoutToCustom = useCallback(() => {
        if (selectedLayoutPresetId !== "custom") {
            setSelectedLayoutPresetId("custom");
        }
   },[selectedLayoutPresetId]);


   const resetState = useCallback(() => {
    setFile(null);
    setFileName('');
    setHeaders([]);
    setFileData([]);
    setColumnMappings([]);
    // Reset output config to default, keeping encoding
    setOutputConfig(prev => ({
        name: 'Configuração Atual',
        format: 'txt',
        fields: [],
        encoding: prev.encoding, // Keep the last selected encoding
    }));
    setSelectedLayoutPresetId("custom"); // Reset layout preset
    // Don't reset predefined fields loaded from storage
    setNewFieldName('');
    setConvertedData('');
    setIsProcessing(false);
    setProcessingMessage('Processando...');
    setActiveTab("upload");
    setShowPreview(false);
     setStaticFieldDialogState({ isOpen: false, isEditing: false, fieldName: '', staticValue: '', length: '', paddingChar: ' ', paddingDirection: 'right' });
     setPredefinedFieldDialogState({ isOpen: false, isEditing: false, fieldName: '', isPersistent: false, comment: '' });
      setDownloadDialogState({ isOpen: false, proposedFilename: '', finalFilename: '' });
      setCalculatedFieldDialogState({ isOpen: false, isEditing: false, fieldName: '', type: '', parameters: {}, requiredInputFields: { parcelasPagas: null, valorRealizado: null }, length: '', paddingChar: ' ', paddingDirection: 'right', dateFormat: '' });
      setConfigManagementDialogState({ isOpen: false, action: null, configName: '', selectedConfigToLoad: null });
    const fileInput = document.getElementById('file-upload') as HTMLInputElement;
    if (fileInput) fileInput.value = '';
     toast({ title: "Pronto", description: "Formulário reiniciado para nova conversão." });
  }, [toast]);


  // --- File Handling ---
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      const allowedTypes = [
          'application/vnd.ms-excel',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'application/vnd.oasis.opendocument.spreadsheet',
          'text/csv',
          // For .ret, MIME type might be generic like 'application/octet-stream' or 'text/plain'
          // So, we'll primarily rely on extension for .ret
      ];
      const fileExtension = selectedFile.name.split('.').pop()?.toLowerCase();

      if (!allowedTypes.includes(selectedFile.type) && fileExtension !== 'csv' && fileExtension !== 'ret') {
        toast({
          title: "Erro",
          description: "Tipo de arquivo inválido. Por favor, selecione XLS, XLSX, ODS, CSV ou RET.",
          variant: "destructive",
        });
        setFile(null);
        setFileName('');
        const fileInput = event.target as HTMLInputElement;
        if(fileInput) fileInput.value = '';
        return;
      }
      setFile(selectedFile);
      setFileName(selectedFile.name);
      setHeaders([]);
      setFileData([]);
      setColumnMappings([]);
      setConvertedData('');
      setOutputConfig(prev => ({ ...prev, name: 'Configuração Atual', fields: [] })); // Reset fields but keep other settings
      setSelectedLayoutPresetId("custom"); // Reset layout on new file
      setActiveTab("mapping");
      processFile(selectedFile);
    }
  };

   // --- Guessing Logic (Moved before processFile) ---
   const guessPredefinedField = useCallback((header: string): string | null => {
      const lowerHeader = header.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Normalize and remove accents
      const guesses: { [key: string]: {id: string, keywords: string[]} } = {
          'matricula': {id: 'matricula', keywords: ['matricula', 'mat', 'registro', 'id func', 'cod func']},
          'cpf': {id: 'cpf', keywords: ['cpf', 'cadastro pessoa fisica']},
          'rg': {id: 'rg', keywords: ['rg', 'identidade', 'registro geral']},
          'nome': {id: 'nome', keywords: ['nome', 'nome completo', 'funcionario', 'colaborador', 'name', 'servidor']},
          'email': {id: 'email', keywords: ['email', 'e-mail', 'correio eletronico', 'contato']},
          'cnpj': {id: 'cnpj', keywords: ['cnpj', 'cadastro nacional pessoa juridica']},
          'regime': {id: 'regime', keywords: ['regime', 'tipo regime']},
          'situacao_usuario': {id: 'situacao_usuario', keywords: ['situacao', 'status', 'situacao usuario', 'situacao do usuario']},
          'categoria': {id: 'categoria', keywords: ['categoria']},
          'secretaria': {id: 'secretaria', keywords: ['secretaria', 'orgao', 'unidade', 'local', 'local de trabalho', 'lotacao', 'departamento', 'setor']}, // Added 'setor' here
          // 'setor' key can be removed if 'secretaria' covers it well enough.
          // {id: 'setor', keywords: ['setor']},
          'margem_bruta': {id: 'margem_bruta', keywords: ['margem bruta', 'valor bruto', 'bruto', 'salario bruto', 'margem']},
          'margem_reservada': {id: 'margem_reservada', keywords: ['margem reservada', 'reservada', 'valor reservado', 'margem cartao', 'margem_cartao', 'margem2']},
          'margem_liquida': {id: 'margem_liquida', keywords: ['margem liquida', 'liquido', 'valor liquido', 'disponivel', 'margem disponivel', 'margem1']},
          'parcelas_pagas': {id: 'parcelas_pagas', keywords: ['parcelas pagas', 'parc pagas', 'qtd parcelas pagas', 'parc']},
          'data_nascimento': {id: 'data_nascimento', keywords: ['data nascimento', 'dt nasc', 'nascimento']},
          'data_admissao': {id: 'data_admissao', keywords: ['data admissao', 'dt adm', 'admissao']},
          'data_fim_contrato': {id: 'data_fim_contrato', keywords: ['data fim contrato', 'dt fim', 'termino contrato']},
          'sinal_margem': {id: 'sinal_margem', keywords: ['sinal margem', 'sinal']},
          'estabelecimento_empresa': {id: 'estabelecimento_empresa', keywords: ['estabelecimento', 'empresa', 'razao social', 'nome fantasia', 'estab']},
          'consignataria_contrato': {id: 'consignataria_contrato', keywords: ['consignataria', 'consignataria contrato', 'empresa contrato']},
          'orgao_filial': {id: 'orgao_filial', keywords: ['orgao filial', 'filial', 'unidade filial', 'orgao pagador', 'org']},
          'verba_rubrica': {id: 'verba_rubrica', keywords: ['verba', 'rubrica', 'cod verba', 'cod rubrica', 'cod_verba']},
          'prazo_total': {id: 'prazo_total', keywords: ['prazo total', 'total parcelas', 'num parcelas', 'prazo']},
          'parcelas_restantes': {id: 'parcelas_restantes', keywords: ['parcelas restantes', 'parc restantes', 'saldo parcelas']},
          'valor_parcela': {id: 'valor_parcela', keywords: ['valor parcela', 'vlr parcela', 'prestacao', 'valor historico', 'valor']},
          'valor_financiado': {id: 'valor_financiado', keywords: ['valor financiado', 'vlr financiado', 'montante']},
          'valor_total': {id: 'valor_total', keywords: ['valor total', 'vlr total', 'total geral']},
          'cet_mensal': {id: 'cet_mensal', keywords: ['cet mensal', 'taxa mes']},
          'cet_anual': {id: 'cet_anual', keywords: ['cet anual', 'taxa ano']},
          'numero_contrato': {id: 'numero_contrato', keywords: ['numero contrato', 'contrato', 'num contrato', 'nro contrato', 'n_contrato']},
          'verba_rubrica_ferias': {id: 'verba_rubrica_ferias', keywords: ['verba ferias', 'rubrica ferias', 'cod verba ferias']},
          'valor_previsto': {id: 'valor_previsto', keywords: ['valor previsto', 'vlr prev', 'previsto']},
          'valor_realizado': {id: 'valor_realizado', keywords: ['valor realizado', 'vlr real', 'realizado']},
          'observacao': {id: 'observacao', keywords: ['observacao', 'obs', 'detalhes', 'motivo']},
          'motivo_critica': {id: 'motivo_critica', keywords: ['motivo critica', 'critica', 'detalhe retorno', 'motivo retorno']},
          'situacao_parcela': {id: 'situacao_parcela', keywords: ['situacao parcela', 'status parcela']},
          'periodo': {id: 'periodo', keywords: ['periodo', 'competencia', 'mes ref']},
          'identificador': {id: 'identificador', keywords: ['identificador', 'id', 'codigo', 'chave']},
          'indice': {id: 'indice', keywords: ['indice', 'fator', 'taxa indice']},
          'tempo_casa': {id: 'tempo_casa', keywords: ['tempo de casa', 'tempo casa', 'antiguidade']},
      };

      const sortedGuessKeys = Object.keys(guesses).sort((a, b) => {
          const aKeywords = guesses[a].keywords.join(' ');
          const bKeywords = guesses[b].keywords.join(' ');
          if (bKeywords.length !== aKeywords.length) {
            return bKeywords.length - aKeywords.length;
          }
          return guesses[b].keywords.length - guesses[a].keywords.length;
      });

      for (const key of sortedGuessKeys) {
          const guess = guesses[key];
          if (guess.keywords.some(keyword => lowerHeader.includes(keyword))) {
              if (predefinedFields.some(pf => pf.id === guess.id)) {
                  return guess.id;
              }
          }
      }
      return null;
  }, [predefinedFields]);

 const guessDataType = useCallback((header: string, sampleData: any): DataType | null => {
      const lowerHeader = header.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      const stringSample = String(sampleData).trim();

       // Priority based on header keywords
      if (lowerHeader.includes('cnpj')) return 'CNPJ';
      if (lowerHeader.includes('cpf')) return 'CPF';
      if (lowerHeader.includes('data') || lowerHeader.includes('date') || lowerHeader.includes('nasc') || lowerHeader.includes('periodo') || lowerHeader.includes('admissao') || lowerHeader.includes('fim contrato')) return 'Data';
      if (lowerHeader.includes('matricula') || lowerHeader.includes('mat') ) return 'Alfanumérico'; // Matrícula is Alfanumérico
      if (lowerHeader.includes('margem') || lowerHeader.includes('valor') || lowerHeader.includes('salario') || lowerHeader.includes('saldo') || lowerHeader.includes('preco') || lowerHeader.includes('brut') || lowerHeader.includes('liquid') || lowerHeader.includes('reservad') || lowerHeader.includes('parcela') || lowerHeader.includes('financiado') || lowerHeader.includes('cet') || lowerHeader.includes('previsto') || lowerHeader.includes('realizado') || lowerHeader.includes('total') ) return 'Numérico';
      if (lowerHeader.includes('cod') || lowerHeader.includes('numero') || lowerHeader.includes('num') || lowerHeader.includes('id') || lowerHeader.includes('prazo') || lowerHeader.includes('restante') ) return 'Inteiro'; // Default to Inteiro if not Matrícula
      if (lowerHeader.includes('rg') || lowerHeader.includes('sinal') || lowerHeader.includes('contrato') || lowerHeader.includes('identificador') || lowerHeader.includes('indice') || lowerHeader.includes('tempo casa')) return 'Alfanumérico';
      if (lowerHeader.includes('idade') || lowerHeader.includes('quant')) return 'Numérico';
      if (lowerHeader.includes('nome') || lowerHeader.includes('descri') || lowerHeader.includes('obs') || lowerHeader.includes('secretaria') || lowerHeader.includes('setor') || lowerHeader.includes('local') || lowerHeader.includes('regime') || lowerHeader.includes('situacao') || lowerHeader.includes('categoria') || lowerHeader.includes('email') || lowerHeader.includes('orgao') || lowerHeader.includes('cargo') || lowerHeader.includes('funcao') || lowerHeader.includes('empresa') || lowerHeader.includes('filial') || lowerHeader.includes('verba') || lowerHeader.includes('rubrica') || lowerHeader.includes('consignataria') || lowerHeader.includes('critica')) return 'Alfanumérico';

      // Guess based on sample data content if header wasn't decisive
       if (stringSample) {
            if (/^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(stringSample) || /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(stringSample) || /^\d{6,8}$/.test(stringSample)) return 'Data';
            if (/^\d{3}\.?\d{3}\.?\d{3}-?\d{2}$/.test(stringSample)) return 'CPF';
            if (/^\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}$/.test(stringSample)) return 'CNPJ';
            if (/[R$]/.test(stringSample) || /[,.]\d{2}$/.test(stringSample) || /^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(stringSample) || /^-?\d+,\d+$/.test(stringSample) ) return 'Numérico';
            if (/^-?\d+$/.test(stringSample)) return 'Inteiro';
            if (/^-?\d+(\.\d+)?$/.test(stringSample)) return 'Numérico';
        }

       if (/[a-zA-Z]/.test(lowerHeader) || (stringSample && /[a-zA-Z]/.test(stringSample))) return 'Alfanumérico';
       if (/^\d+$/.test(lowerHeader)) return 'Inteiro';

      return 'Alfanumérico';
  }, []);


 const processFile = useCallback(async (fileToProcess: File) => {
     if (!fileToProcess) return;
     setIsProcessing(true);
     setProcessingMessage('Lendo arquivo...');
     setHeaders([]);
     setFileData([]);
     setColumnMappings([]);
     setConvertedData('');
     setOutputConfig(prev => ({ ...prev, name: 'Configuração Atual', fields: [] }));
     setSelectedLayoutPresetId("custom");

     let extractedHeaders: string[] = [];
     let extractedData: any[] = [];
     const fileExtension = fileToProcess.name.split('.').pop()?.toLowerCase();

     try {
         if (fileExtension === 'csv') {
             setProcessingMessage('Processando arquivo CSV...');
             const fileText = await fileToProcess.text();
             const lines = fileText.split(/\r?\n/).filter(line => line.trim() !== '');
             if (lines.length === 0) throw new Error("Arquivo CSV vazio ou inválido.");

             // Detect delimiter
             let delimiter = ';'; // Default
             const firstLine = lines[0];
             if (firstLine.includes('|')) delimiter = '|';
             else if (firstLine.includes('\t')) delimiter = '\t';
             else if (firstLine.includes(',')) delimiter = ','; // Add comma as a common delimiter

             extractedHeaders = lines[0].split(delimiter).map(h => h.trim().replace(/^["']|["']$/g, ''));
             extractedData = lines.slice(1).map(rowString => {
                 const values = rowString.split(delimiter);
                 const rowData: { [key: string]: any } = {};
                 extractedHeaders.forEach((header, index) => {
                     rowData[header] = (values[index] ?? '').trim().replace(/^["']|["']$/g, '');
                 });
                 return rowData;
             });
              toast({ title: "Sucesso", description: `Arquivo CSV ${fileToProcess.name} processado.` });

         } else if (fileExtension === 'ret') {
             setProcessingMessage('Processando arquivo CNAB (.ret)...');
             const arrayBuffer = await fileToProcess.arrayBuffer();
             let text;
             try {
                 text = iconv.decode(Buffer.from(arrayBuffer), 'ISO-8859-1');
             } catch (e) {
                 console.warn("Falha ao decodificar .RET como ISO-8859-1, tentando UTF-8...");
                 text = iconv.decode(Buffer.from(arrayBuffer), 'UTF-8');
             }
             const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
             if (lines.length === 0) throw new Error("Arquivo .RET CNAB vazio ou inválido.");

             extractedHeaders = ['LinhaCNAB'];
             extractedData = lines.map(line => ({ 'LinhaCNAB': line }));
             toast({ title: "Aviso", description: `Processamento básico de CNAB .RET: ${fileToProcess.name}. Cada linha é um campo.`, variant: "default", duration: 7000 });

         } else if (fileToProcess.type.includes('spreadsheet') || fileToProcess.type.includes('excel') || fileToProcess.name.endsWith('.ods')) {
             setProcessingMessage('Processando planilha...');
             const data = await fileToProcess.arrayBuffer();
             const workbook = XLSX.read(data, { type: 'array' });
             const sheetName = workbook.SheetNames[0];
             const worksheet = workbook.Sheets[sheetName];
             const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });

             if (jsonData.length > 0) {
                 extractedHeaders = jsonData[0].map(String);
                 extractedData = jsonData.slice(1).map(row => {
                     const rowData: { [key: string]: any } = {};
                     extractedHeaders.forEach((header, index) => {
                         rowData[header] = row[index] ?? '';
                     });
                     return rowData;
                 });
             }
         } else {
             throw new Error("Tipo de arquivo não suportado para processamento.");
         }


         if (extractedHeaders.length === 0 && extractedData.length > 0) {
            if (fileExtension !== 'ret') { // For RET, we explicitly set one header
                extractedHeaders = Object.keys(extractedData[0]).map((key, i) => `Coluna ${i + 1}`);
                toast({ title: "Aviso", description: "Cabeçalhos não encontrados, usando 'Coluna 1', 'Coluna 2', etc.", variant: "default" });
            }
         } else if (extractedHeaders.length === 0) {
              throw new Error("Não foi possível extrair cabeçalhos ou dados do arquivo.");
          }

         setHeaders(extractedHeaders);
         setFileData(extractedData);

         const newColumnMappings = extractedHeaders.map(header => {
             const lowerHeader = header.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
             let guessedFieldId = guessPredefinedField(header);
             let guessedType = guessDataType(header, extractedData.length > 0 ? extractedData[0][header] : '');
             let removeMask = !!guessedFieldId && ['cpf', 'rg', 'cnpj'].includes(guessedFieldId) || ['Data', 'Numérico', 'Inteiro', 'CPF', 'CNPJ'].includes(guessedType ?? '');

             // Specific overrides based on user request
             const valorFieldKeywords = ["valor parcela", "valor financiado", "valor total", "valor previsto", "valor realizado"];
             const valorFieldIdsFromGuess = ['valor_parcela', 'valor_financiado', 'valor_total', 'valor_previsto', 'valor_realizado'];
             const localTrabalhoKeywords = ['local', 'setor', 'local de trabalho', 'lotacao', 'departamento', 'secretaria'];


             if (lowerHeader === 'margem') {
                 guessedFieldId = 'margem_bruta';
                 guessedType = 'Numérico';
                 removeMask = false;
             } else if (guessedFieldId === 'margem_bruta') {
                 guessedType = 'Numérico';
                 removeMask = false;
             } else if (localTrabalhoKeywords.some(kw => lowerHeader.includes(kw))) {
                 guessedFieldId = 'secretaria'; // ID for 'Local de Trabalho'
                 guessedType = 'Alfanumérico';
             } else if (guessedFieldId === 'secretaria') { // ID for 'Local de Trabalho'
                 guessedType = 'Alfanumérico';
             } else if (lowerHeader === 'matrícula' || lowerHeader === 'matricula') {
                 guessedFieldId = 'matricula';
                 guessedType = 'Alfanumérico';
             } else if (valorFieldKeywords.some(kw => lowerHeader.includes(kw)) || valorFieldIdsFromGuess.includes(guessedFieldId || '')) {
                 guessedType = 'Numérico';
                 removeMask = false;
                 // Ensure correct field ID if matched by keyword but guessedFieldId doesn't match a valor field
                 if (valorFieldKeywords.some(kw => lowerHeader.includes(kw)) && !valorFieldIdsFromGuess.includes(guessedFieldId || '')) {
                    if (lowerHeader.includes('valor parcela')) guessedFieldId = 'valor_parcela';
                    else if (lowerHeader.includes('valor financiado')) guessedFieldId = 'valor_financiado';
                    else if (lowerHeader.includes('valor total')) guessedFieldId = 'valor_total';
                    else if (lowerHeader.includes('valor previsto')) guessedFieldId = 'valor_previsto';
                    else if (lowerHeader.includes('valor realizado')) guessedFieldId = 'valor_realizado';
                 }
             }

             return {
                 originalHeader: header,
                 mappedField: guessedFieldId,
                 dataType: guessedType,
                 length: null,
                 removeMask: removeMask,
             };
         });
         setColumnMappings(newColumnMappings);

         toast({ title: "Sucesso", description: `Arquivo ${fileToProcess.name} processado (${extractedData.length} linhas). Verifique o mapeamento.` });
         setActiveTab("mapping");

     } catch (error: any) {
         console.error("Erro ao processar arquivo:", error);
         toast({
             title: "Erro ao Processar Arquivo",
             description: error.message || "Ocorreu um erro inesperado.",
             variant: "destructive",
         });
         setActiveTab("upload");
         setHeaders([]);
         setFileData([]);
         setColumnMappings([]);
     } finally {
         setIsProcessing(false);
         setProcessingMessage('Processando...');
     }
 }, [toast, guessPredefinedField, guessDataType, predefinedFields]);




  // --- Mapping ---
  const handleMappingChange = (index: number, field: keyof ColumnMapping, value: any) => {
    setColumnMappings(prev => {
      const newMappings = [...prev];
      const currentMapping = { ...newMappings[index] };
      let actualValue = value === NONE_VALUE_PLACEHOLDER ? null : value;

      if (field === 'dataType') {
         (currentMapping[field] as any) = actualValue;
         if (newMappings[index].dataType === 'Alfanumérico' && actualValue !== 'Alfanumérico') {
             currentMapping.length = null;
         }
         currentMapping.removeMask = ['CPF', 'RG', 'CNPJ', 'Data', 'Numérico', 'Inteiro'].includes(actualValue ?? '');

         // Special override for margem_bruta or valor fields if type is changed away from Numérico manually
         const isValorField = ['margem_bruta', 'valor_parcela', 'valor_financiado', 'valor_total', 'valor_previsto', 'valor_realizado'].includes(currentMapping.mappedField || '');
         if (isValorField && actualValue !== 'Numérico') {
            currentMapping.removeMask = ['CPF', 'RG', 'CNPJ', 'Data', 'Numérico', 'Inteiro'].includes(actualValue ?? ''); // revert to default
         } else if (isValorField && actualValue === 'Numérico') {
            currentMapping.removeMask = false;
         }
         if (currentMapping.mappedField === 'matricula' && actualValue !== 'Alfanumérico') {
            // allow changing if user insists, but it defaults to Alfanumérico
         }


       } else if (field === 'length') {
           const numValue = parseInt(value, 10);
           currentMapping.length = (currentMapping.dataType === 'Alfanumérico' && !isNaN(numValue) && numValue > 0) ? numValue : null;
       } else if (field === 'removeMask') {
            const isValorField = ['margem_bruta', 'valor_parcela', 'valor_financiado', 'valor_total', 'valor_previsto', 'valor_realizado'].includes(currentMapping.mappedField || '');
            if (isValorField && currentMapping.dataType === 'Numérico') {
                 currentMapping.removeMask = false; // Enforce no mask removal
                 if (Boolean(value) === true) {
                     toast({title: "Aviso", description: `Remoção de máscara é desabilitada para '${predefinedFields.find(pf => pf.id === currentMapping.mappedField)?.name || currentMapping.mappedField}' do tipo 'Numérico'.`, variant: "default"});
                 }
            } else {
                currentMapping.removeMask = Boolean(value);
            }
       } else {
          (currentMapping[field] as any) = actualValue;
            if (field === 'mappedField' && actualValue) {
                 const predefined = predefinedFields.find(pf => pf.id === actualValue);
                  const sampleData = fileData.length > 0 ? fileData[0][currentMapping.originalHeader] : '';
                 let guessedType = predefined ? guessDataType(predefined.name, sampleData) : guessDataType(currentMapping.originalHeader, sampleData);
                 let removeMask = ['CPF', 'RG', 'CNPJ', 'Data', 'Numérico', 'Inteiro'].includes(guessedType ?? '');

                if (actualValue === 'matricula') {
                    guessedType = 'Alfanumérico';
                } else if (['margem_bruta', 'valor_parcela', 'valor_financiado', 'valor_total', 'valor_previsto', 'valor_realizado'].includes(actualValue)) {
                    guessedType = 'Numérico';
                    removeMask = false;
                 } else if (actualValue === 'secretaria') { // ID for 'Local de Trabalho'
                    guessedType = 'Alfanumérico';
                 }

                 currentMapping.dataType = guessedType;
                 currentMapping.removeMask = removeMask;
                  if (guessedType !== 'Alfanumérico') {
                      currentMapping.length = null;
                  }
            } else if (field === 'mappedField' && !actualValue) { // If unmapping
                currentMapping.dataType = null; // Clear data type
                currentMapping.removeMask = false; // Reset removeMask
            }
       }

      newMappings[index] = currentMapping;
      return newMappings;
    });
  };



  // --- Predefined Fields ---
   const openAddPredefinedFieldDialog = () => {
        setPredefinedFieldDialogState({
            isOpen: true,
            isEditing: false,
            fieldName: '',
            isPersistent: false,
            comment: '',
        });
    };

    const openEditPredefinedFieldDialog = (field: PredefinedField) => {
        setPredefinedFieldDialogState({
            isOpen: true,
            isEditing: true,
            fieldId: field.id,
            fieldName: field.name,
            isPersistent: field.isPersistent || false,
            comment: field.comment || '',
        });
    };

    const handlePredefinedFieldDialogChange = (field: keyof PredefinedFieldDialogState, value: any) => {
         setPredefinedFieldDialogState(prev => ({
             ...prev,
             [field]: value
         }));
     };

   const savePredefinedField = () => {
        const { isEditing, fieldId, fieldName, isPersistent, comment } = predefinedFieldDialogState;
        const trimmedName = fieldName.trim();

        if (!trimmedName) {
            toast({ title: "Erro", description: "Nome do campo não pode ser vazio.", variant: "destructive" });
            return;
        }

         const newId = isEditing ? fieldId! : trimmedName.toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');

          if (!newId) {
             toast({ title: "Erro", description: "Nome do campo inválido para gerar um ID.", variant: "destructive" });
             return;
         }

         if (!isEditing && predefinedFields.some(f => f.id === newId)) {
             toast({ title: "Erro", description: `Já existe um campo com o ID gerado "${newId}". Escolha um nome diferente.`, variant: "destructive" });
             return;
         }
          if (predefinedFields.some(f => f.name.toLowerCase() === trimmedName.toLowerCase() && f.id !== fieldId)) {
              toast({ title: "Erro", description: `Já existe um campo com o nome "${trimmedName}". Escolha um nome diferente.`, variant: "destructive" });
              return;
          }

          let updatedFields: PredefinedField[];
          let fieldDescription = `Campo "${trimmedName}"`;
          let fieldToUpdateOrAdd: PredefinedField;

          if (isEditing) {
              const originalField = predefinedFields.find(f => f.id === fieldId);
              if (!originalField) return;

               fieldToUpdateOrAdd = {
                  ...originalField,
                  name: trimmedName,
                  comment: comment || '',
                  isPersistent: isPersistent,
                  group: originalField.isCore ? originalField.group : (isPersistent ? 'Principal Personalizado' : 'Opcional Personalizado')
               };

              updatedFields = predefinedFields.map(f =>
                  f.id === fieldId ? fieldToUpdateOrAdd : f
              );
              fieldDescription += ` atualizado (${isPersistent ? 'Principal' : 'Opcional'}).`;

          } else {
              fieldToUpdateOrAdd = {
                  id: newId,
                  name: trimmedName,
                  comment: comment || '',
                  isCore: false,
                  isPersistent: isPersistent,
                  group: isPersistent ? 'Principal Personalizado' : 'Opcional Personalizado'
              };
              updatedFields = [...predefinedFields, fieldToUpdateOrAdd];
              fieldDescription += ` adicionado com ID "${newId}" (${isPersistent ? 'Principal' : 'Opcional'}).`;
          }

        setPredefinedFields(updatedFields.sort((a,b) => a.name.localeCompare(b.name)));
        saveCustomPredefinedFields(updatedFields);
        setPredefinedFieldDialogState({ isOpen: false, isEditing: false, fieldName: '', isPersistent: false, comment: '' });
        toast({ title: "Sucesso", description: fieldDescription });
    };


  const removePredefinedField = (idToRemove: string) => {
    const fieldToRemove = predefinedFields.find(f => f.id === idToRemove);
    if (!fieldToRemove) return;

     const updatedFields = predefinedFields.filter(f => f.id !== idToRemove);
    setPredefinedFields(updatedFields.sort((a,b) => a.name.localeCompare(b.name)));

    setColumnMappings(prev => prev.map(m => m.mappedField === idToRemove ? { ...m, mappedField: null, dataType: null, removeMask: false } : m));
    setOutputConfig(prev => ({
      ...prev,
       fields: prev.fields
            .filter(f => {
                 if (f.isStatic) return true;
                 if (f.isCalculated) {
                     return !f.requiredInputFields.includes(idToRemove);
                 }
                 return f.mappedField !== idToRemove;
            })
            .map((f, idx) => ({ ...f, order: idx })),
    }));

     saveCustomPredefinedFields(updatedFields);
     resetSelectedLayoutToCustom();

    toast({ title: "Sucesso", description: `Campo "${fieldToRemove.name}" removido.` });
  };


  // --- Output Configuration ---
   const handleOutputFormatChange = (value: OutputFormat) => {
      setOutputConfig(prev => {
          const newFields = prev.fields.map(f => ({
              ...f,
              delimiter: value === 'csv' ? (prev.delimiter || '|') : undefined,
               length: value === 'txt' ? (f.length ?? (f.isStatic ? (f.staticValue?.length || 10) : f.isCalculated ? 10 : 10)) : f.length,
               paddingChar: value === 'txt' ? (f.paddingChar ?? getDefaultPaddingChar(f, columnMappings)) : f.paddingChar,
               paddingDirection: value === 'txt' ? (f.paddingDirection ?? getDefaultPaddingDirection(f, columnMappings)) : f.paddingDirection,
          }));
          return {
              ...prev,
              format: value,
              delimiter: value === 'csv' ? (prev.delimiter || '|') : undefined,
              fields: newFields
          };
      });
      if (value !== 'txt' || (value === 'txt' && selectedLayoutPresetId !== "custom")) {
          resetSelectedLayoutToCustom();
      }
  };

  const handleDelimiterChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setOutputConfig(prev => ({ ...prev, delimiter: event.target.value }));
    resetSelectedLayoutToCustom();
  };

 const handleOutputFieldChange = (id: string, field: keyof OutputFieldConfig, value: any) => {
    setOutputConfig(prev => {
        const newFields = prev.fields.map(f => {
            if (f.id === id) {
                const updatedField = { ...f };
                let actualValue = value === NONE_VALUE_PLACEHOLDER ? null : value;

                if (field === 'mappedField' && !updatedField.isStatic && !updatedField.isCalculated) {
                        updatedField.mappedField = actualValue;
                        const correspondingMapping = columnMappings.find(cm => cm.mappedField === actualValue);
                        const dataType = correspondingMapping?.dataType ?? null;

                         if (prev.format === 'txt') {
                            updatedField.length = updatedField.length ?? (correspondingMapping?.dataType === 'Alfanumérico' ? correspondingMapping.length : 10) ?? 10;
                             updatedField.paddingChar = updatedField.paddingChar ?? getDefaultPaddingChar(updatedField, columnMappings);
                             updatedField.paddingDirection = updatedField.paddingDirection ?? getDefaultPaddingDirection(updatedField, columnMappings);
                         }

                        if (dataType === 'Data') {
                            updatedField.dateFormat = updatedField.dateFormat ?? 'DDMMYYYY'; // Default to DDMMYYYY for new date fields
                        } else {
                            delete updatedField.dateFormat;
                        }
                } else if (field === 'length') {
                    const numValue = parseInt(value, 10);
                    if (prev.format === 'txt') {
                        updatedField.length = isNaN(numValue) || numValue <= 0 ? undefined : numValue;
                        updatedField.paddingChar = updatedField.paddingChar ?? getDefaultPaddingChar(updatedField, columnMappings);
                        updatedField.paddingDirection = updatedField.paddingDirection ?? getDefaultPaddingDirection(updatedField, columnMappings);
                    } else {
                         delete updatedField.length;
                    }
                } else if (field === 'order') {
                     console.warn("Changing order directly is disabled. Use move buttons.");
                    return updatedField;
                } else if (field === 'paddingChar') {
                    if (prev.format === 'txt') {
                         updatedField.paddingChar = String(value).slice(0, 1);
                     } else {
                        delete updatedField.paddingChar;
                    }
                } else if (field === 'paddingDirection') {
                    if (prev.format === 'txt') {
                        updatedField.paddingDirection = value as PaddingDirection;
                    } else {
                        delete updatedField.paddingDirection;
                    }
                } else if (field === 'dateFormat') {
                     const currentDataType = getOutputFieldDataType(updatedField);
                      if (currentDataType === 'Data') {
                          updatedField.dateFormat = value === NONE_VALUE_PLACEHOLDER ? undefined : value as DateFormat;
                      } else {
                          delete updatedField.dateFormat;
                      }
                }
                else {
                     if (field === 'staticValue' && updatedField.isStatic) {
                         updatedField.staticValue = actualValue;
                         if (prev.format === 'txt') {
                              updatedField.length = updatedField.length ?? updatedField.staticValue.length ?? 10;
                              updatedField.paddingChar = updatedField.paddingChar ?? getDefaultPaddingChar(updatedField, columnMappings);
                              updatedField.paddingDirection = updatedField.paddingDirection ?? getDefaultPaddingDirection(updatedField, columnMappings);
                          }
                     } else if (field === 'fieldName' && (updatedField.isStatic || updatedField.isCalculated)) {
                         updatedField.fieldName = actualValue;
                     }
                }
                return updatedField;
            }
            return f;
        });

        return { ...prev, fields: newFields };
    });
    resetSelectedLayoutToCustom();
};


  const addOutputField = () => {
    const availableMappedFields = columnMappings
        .filter(m => m.mappedField !== null && !outputConfig.fields.some(of => !of.isStatic && !of.isCalculated && of.mappedField === m.mappedField))
        .map(m => m.mappedField);

    if (availableMappedFields.length === 0) {
        toast({ title: "Aviso", description: "Não há mais campos mapeados disponíveis para adicionar.", variant: "default"});
        return;
    }

    const maxOrder = outputConfig.fields.length > 0 ? Math.max(...outputConfig.fields.map(f => f.order)) : -1;
    const newFieldId = availableMappedFields[0]!;
    const correspondingMapping = columnMappings.find(cm => cm.mappedField === newFieldId);
    const dataType = correspondingMapping?.dataType ?? null;
    const defaultLength = (outputConfig.format === 'txt' && dataType === 'Alfanumérico') ? (correspondingMapping?.length ?? 10) : (outputConfig.format === 'txt' ? 10 : undefined);

    const newOutputField: OutputFieldConfig = {
        id: `mapped-${newFieldId}-${Date.now()}`,
        isStatic: false,
        isCalculated: false,
        mappedField: newFieldId,
        order: maxOrder + 1,
        ...(outputConfig.format === 'txt' && {
             length: defaultLength,
             paddingChar: getDefaultPaddingChar({isStatic: false, isCalculated: false, mappedField: newFieldId, id: '', order: 0 }, columnMappings),
             paddingDirection: getDefaultPaddingDirection({isStatic: false, isCalculated: false, mappedField: newFieldId, id: '', order: 0 }, columnMappings),
         }),
        dateFormat: dataType === 'Data' ? 'DDMMYYYY' : undefined, // Default to DDMMYYYY for new date fields
    };

    setOutputConfig(prev => ({
        ...prev,
        fields: [...prev.fields, newOutputField].sort((a, b) => a.order - b.order)
    }));
    resetSelectedLayoutToCustom();
};


  const removeOutputField = (idToRemove: string) => {
     setOutputConfig(prev => {
         const newFields = prev.fields.filter(f => f.id !== idToRemove);
         const reorderedFields = newFields.sort((a, b) => a.order - b.order).map((f, idx) => ({ ...f, order: idx }));
         return {
             ...prev,
             fields: reorderedFields,
         };
     });
     resetSelectedLayoutToCustom();
   };

   const moveField = (id: string, direction: 'up' | 'down') => {
       setOutputConfig(prev => {
           const fields = [...prev.fields];
           const currentIndex = fields.findIndex(f => f.id === id);

           if (currentIndex === -1) return prev;

           const targetIndex = direction === 'up' ? currentIndex - 1 : currentIndex + 1;

           if (targetIndex < 0 || targetIndex >= fields.length) return prev;

           const currentOrder = fields[currentIndex].order;
           fields[currentIndex].order = fields[targetIndex].order;
           fields[targetIndex].order = currentOrder;

           fields.sort((a, b) => a.order - b.order);
            const renumberedFields = fields.map((f, idx) => ({ ...f, order: idx }));

           return { ...prev, fields: renumberedFields };
       });
       resetSelectedLayoutToCustom();
   };


    const openAddStaticFieldDialog = () => {
        setStaticFieldDialogState({
            isOpen: true,
            isEditing: false,
            fieldName: '',
            staticValue: '',
            length: '',
            paddingChar: ' ',
            paddingDirection: 'right',
        });
    };

    const openEditStaticFieldDialog = (field: OutputFieldConfig) => {
        if (!field.isStatic) return;
        setStaticFieldDialogState({
            isOpen: true,
            isEditing: true,
            fieldId: field.id,
            fieldName: field.fieldName,
            staticValue: field.staticValue,
            length: String(field.length ?? ''),
            paddingChar: field.paddingChar ?? getDefaultPaddingChar(field, columnMappings),
            paddingDirection: field.paddingDirection ?? getDefaultPaddingDirection(field, columnMappings),
        });
    };

    const handleStaticFieldDialogChange = (field: keyof StaticFieldDialogState, value: any) => {
        setStaticFieldDialogState(prev => ({
            ...prev,
            [field]: value
        }));
    };

    const saveStaticField = () => {
        const { isEditing, fieldId, fieldName, staticValue, length, paddingChar, paddingDirection } = staticFieldDialogState;
        const len = parseInt(length, 10);
        const isTxtFormat = outputConfig.format === 'txt';

        if (!fieldName.trim()) {
            toast({ title: "Erro", description: "Nome do Campo Estático não pode ser vazio.", variant: "destructive" });
            return;
        }
        if (isTxtFormat && (isNaN(len) || len <= 0)) {
            toast({ title: "Erro", description: "Tamanho deve ser um número positivo para formato TXT.", variant: "destructive" });
            return;
        }
         if (isTxtFormat && (!paddingChar || paddingChar.length !== 1)) {
            toast({ title: "Erro", description: "Caractere de Preenchimento deve ser um único caractere para TXT.", variant: "destructive" });
            return;
        }

         let staticFieldBase: Omit<OutputFieldConfig, 'id' | 'order' | 'isStaticFromPreset'> & { isStatic: true, isCalculated: false } = {
             isStatic: true,
             isCalculated: false,
             fieldName: fieldName.trim(),
             staticValue: staticValue,
         };

         let staticField: OutputFieldConfig = {
             ...staticFieldBase,
             id: isEditing && fieldId ? fieldId : `static-${Date.now()}`,
             order: 0,
             isStaticFromPreset: false,
             ...(isTxtFormat && {
                length: len,
                paddingChar: paddingChar,
                paddingDirection: paddingDirection,
            }),
         };


        setOutputConfig(prev => {
            let newFields;
            if (isEditing) {
                 const existingFieldIndex = prev.fields.findIndex(f => f.id === fieldId);
                 if (existingFieldIndex === -1) return prev;
                 newFields = [...prev.fields];
                  const updatedStaticField = {
                     ...staticField,
                     order: prev.fields[existingFieldIndex].order,
                     isStaticFromPreset: prev.fields[existingFieldIndex].isStaticFromPreset, // Preserve this flag
                      ...(prev.format === 'txt' && {
                          length: len,
                          paddingChar: paddingChar,
                          paddingDirection: paddingDirection,
                      }),
                     ...(prev.format !== 'txt' && {
                          length: undefined,
                          paddingChar: undefined,
                          paddingDirection: undefined,
                      })
                  };

                 newFields[existingFieldIndex] = updatedStaticField;

            } else {
                 const maxOrder = prev.fields.length > 0 ? Math.max(...prev.fields.map(f => f.order)) : -1;
                 staticField.order = maxOrder + 1;
                 newFields = [...prev.fields, staticField];
            }
             newFields.sort((a, b) => a.order - b.order);
             const reorderedFields = newFields.map((f, idx) => ({ ...f, order: idx }));

            return { ...prev, fields: reorderedFields };
        });

        setStaticFieldDialogState({ ...staticFieldDialogState, isOpen: false });
        toast({ title: "Sucesso", description: `Campo estático "${fieldName.trim()}" ${isEditing ? 'atualizado' : 'adicionado'}.` });
        resetSelectedLayoutToCustom();
    };

    const openAddCalculatedFieldDialog = () => {
        setCalculatedFieldDialogState({
            isOpen: true,
            isEditing: false,
            fieldName: '',
            type: '',
            parameters: {},
            requiredInputFields: { parcelasPagas: null, valorRealizado: null },
            length: '',
            paddingChar: ' ',
            paddingDirection: 'right',
            dateFormat: '',
        });
    };

    const openEditCalculatedFieldDialog = (field: OutputFieldConfig) => {
        if (!field.isCalculated) return;

        const requiredFieldsState: CalculatedFieldDialogState['requiredInputFields'] = { parcelasPagas: null, valorRealizado: null };
        if (field.type === 'CalculateStartDate' && field.requiredInputFields.length > 0) {
            requiredFieldsState.parcelasPagas = field.requiredInputFields[0] ?? null;
        } else if (field.type === 'CalculateSituacaoRetorno' && field.requiredInputFields.length > 0) {
            requiredFieldsState.valorRealizado = field.requiredInputFields[0] ?? null;
        }


        setCalculatedFieldDialogState({
            isOpen: true,
            isEditing: true,
            fieldId: field.id,
            fieldName: field.fieldName,
            type: field.type,
            parameters: { ...field.parameters },
            requiredInputFields: requiredFieldsState,
            length: String(field.length ?? ''),
            paddingChar: field.paddingChar ?? getDefaultPaddingChar(field, columnMappings),
            paddingDirection: field.paddingDirection ?? getDefaultPaddingDirection(field, columnMappings),
            dateFormat: field.dateFormat ?? '',
        });
    };

     const handleCalculatedFieldDialogChange = (
        field: keyof CalculatedFieldDialogState | `parameters.${string}` | `requiredInputFields.${string}`,
        value: any
    ) => {
        setCalculatedFieldDialogState(prev => {
            const newState = { ...prev };
            if (field.startsWith('parameters.')) {
                const paramKey = field.split('.')[1];
                newState.parameters = { ...newState.parameters, [paramKey]: value };
            } else if (field.startsWith('requiredInputFields.')) {
                const inputKey = field.split('.')[1] as keyof CalculatedFieldDialogState['requiredInputFields'];
                newState.requiredInputFields = {
                    ...newState.requiredInputFields,
                    [inputKey]: value === NONE_VALUE_PLACEHOLDER ? null : value
                };
            } else {
                 if (field === 'length') {
                     newState.length = value;
                 } else if (field === 'dateFormat') {
                     newState.dateFormat = value === NONE_VALUE_PLACEHOLDER ? '' : value;
                 } else if (field === 'type') {
                     newState.type = value === NONE_VALUE_PLACEHOLDER ? '' : value as CalculatedFieldType | '';
                      newState.parameters = {};
                      newState.requiredInputFields = { parcelasPagas: null, valorRealizado: null };
                      newState.dateFormat = '';
                      if (value === 'CalculateStartDate') {
                          newState.requiredInputFields = { parcelasPagas: null, valorRealizado: null };
                          newState.dateFormat = 'DDMMYYYY'; // Default to DDMMYYYY
                          newState.paddingDirection = 'left';
                          newState.paddingChar = '0';
                          newState.length = '8';
                      } else if (value === 'CalculateSituacaoRetorno') {
                          newState.requiredInputFields = { parcelasPagas: null, valorRealizado: null };
                          newState.paddingDirection = 'right';
                          newState.paddingChar = ' ';
                          newState.length = '1';
                      } else if (value === 'FormatPeriodMMAAAA') {
                           newState.paddingDirection = 'left';
                           newState.paddingChar = '0';
                           newState.length = '6';
                      }
                      else {
                          newState.length = '';
                          newState.paddingChar = ' ';
                          newState.paddingDirection = 'right';
                      }
                 }
                 else {
                     (newState as any)[field] = value;
                 }
            }
            return newState;
        });
    };

     const saveCalculatedField = () => {
        const { isEditing, fieldId, fieldName, type, parameters, requiredInputFields, length, paddingChar, paddingDirection, dateFormat } = calculatedFieldDialogState;
        const len = parseInt(length, 10);
        const isTxtFormat = outputConfig.format === 'txt';

        if (!fieldName.trim()) {
            toast({ title: "Erro", description: "Nome do Campo Calculado não pode ser vazio.", variant: "destructive" });
            return;
        }
         if (!type) {
             toast({ title: "Erro", description: "Selecione o Tipo de Cálculo.", variant: "destructive" });
             return;
         }
        if (isTxtFormat && (isNaN(len) || len <= 0)) {
            toast({ title: "Erro", description: "Tamanho deve ser um número positivo para formato TXT.", variant: "destructive" });
            return;
        }
         if (isTxtFormat && (!paddingChar || paddingChar.length !== 1)) {
            toast({ title: "Erro", description: "Caractere de Preenchimento deve ser um único caractere para TXT.", variant: "destructive" });
            return;
        }
         if (type === 'CalculateStartDate') {
             if (!parameters.period) {
                 toast({ title: "Erro", description: "Informe o Período Atual (DD/MM/AAAA).", variant: "destructive" });
                 return;
             }
             if (!/^\d{2}\/\d{2}\/\d{4}$/.test(parameters.period) || !dateFnsIsValid(parse(parameters.period, 'dd/MM/yyyy', new Date()))) {
                  toast({ title: "Erro", description: "Formato inválido para Período Atual. Use DD/MM/AAAA.", variant: "destructive" });
                 return;
             }
             if (!requiredInputFields.parcelasPagas) {
                 toast({ title: "Erro", description: "Selecione o campo mapeado para 'Parcelas Pagas'.", variant: "destructive" });
                 return;
             }
             if (!dateFormat) {
                 toast({ title: "Erro", description: "Selecione um Formato de Data para a saída.", variant: "destructive" });
                 return;
             }
         } else if (type === 'CalculateSituacaoRetorno') {
            if (!requiredInputFields.valorRealizado) {
                toast({ title: "Erro", description: "Selecione o campo mapeado para 'Valor Realizado'.", variant: "destructive" });
                return;
            }
         }


        const requiredInputsArray: string[] = [];
        if (type === 'CalculateStartDate' && requiredInputFields.parcelasPagas) {
            requiredInputsArray.push(requiredInputFields.parcelasPagas);
        } else if (type === 'CalculateSituacaoRetorno' && requiredInputFields.valorRealizado) {
            requiredInputsArray.push(requiredInputFields.valorRealizado);
        } else if (type === 'FormatPeriodMMAAAA' && parameters.period) {
            // For FormatPeriodMMAAAA, the input is the 'Período' field if mapped.
            // However, the current setup uses parameters.period for CalculateStartDate.
            // Let's adjust FormatPeriodMMAAAA to also use a mapped 'Período' field or fallback to current date.
            // For now, we assume it might need a 'periodo' mapped field. This will be handled in `calculateFieldValue`.
            // The `requiredInputFields` structure is for mapped fields, so we'll leave it empty for now for this type.
        }


         const calculatedFieldBase: Omit<OutputFieldConfig, 'id' | 'order' | 'isStaticFromPreset'> & { isStatic: false, isCalculated: true } = {
             isStatic: false,
             isCalculated: true,
             fieldName: fieldName.trim(),
             type: type,
             requiredInputFields: requiredInputsArray,
             parameters: { ...parameters },
             ...( (type === 'CalculateStartDate' || type === 'FormatPeriodMMAAAA') && { dateFormat: dateFormat as DateFormat }),
         };

         let calculatedField: OutputFieldConfig = {
              ...calculatedFieldBase,
              id: isEditing && fieldId ? fieldId : `calc-${type}-${Date.now()}`,
              order: 0,
              isStaticFromPreset: false,
             ...(isTxtFormat && {
                 length: len,
                 paddingChar: paddingChar,
                 paddingDirection: paddingDirection,
             }),
         };


        setOutputConfig(prev => {
            let newFields;
            if (isEditing) {
                const existingFieldIndex = prev.fields.findIndex(f => f.id === fieldId);
                if (existingFieldIndex === -1) return prev;
                newFields = [...prev.fields];
                 const updatedCalcField = {
                      ...calculatedField,
                      order: prev.fields[existingFieldIndex].order,
                      isStaticFromPreset: prev.fields[existingFieldIndex].isStaticFromPreset, // Preserve
                      ...(prev.format === 'txt' && {
                          length: len,
                          paddingChar: paddingChar,
                          paddingDirection: paddingDirection,
                      }),
                      ...(prev.format !== 'txt' && {
                          length: undefined,
                          paddingChar: undefined,
                          paddingDirection: undefined,
                      }),
                     ...( (type === 'CalculateStartDate' || type === 'FormatPeriodMMAAAA') && { dateFormat: dateFormat as DateFormat }),
                      ...( (type !== 'CalculateStartDate' && type !== 'FormatPeriodMMAAAA') && { dateFormat: undefined }),
                 };

                newFields[existingFieldIndex] = updatedCalcField;
            } else {
                const maxOrder = prev.fields.length > 0 ? Math.max(...prev.fields.map(f => f.order)) : -1;
                calculatedField.order = maxOrder + 1;
                newFields = [...prev.fields, calculatedField];
            }
            newFields.sort((a, b) => a.order - b.order);
            const reorderedFields = newFields.map((f, idx) => ({ ...f, order: idx }));

            return { ...prev, fields: reorderedFields };
        });

        setCalculatedFieldDialogState({ ...calculatedFieldDialogState, isOpen: false });
        toast({ title: "Sucesso", description: `Campo calculado "${fieldName.trim()}" ${isEditing ? 'atualizado' : 'adicionado'}.` });
        resetSelectedLayoutToCustom();
    };


   useEffect(() => {
        if (columnMappings.length === 0 && fileData.length === 0 && outputConfig.fields.every(f => !f.isStatic && !f.isCalculated)) {
             if(outputConfig.fields.length === 0) return;
        }
        if (selectedLayoutPresetId !== "custom" && outputConfig.format === 'txt') return;


       setOutputConfig(prevConfig => {
           const existingFieldsMap = new Map(prevConfig.fields.map(f => [f.id, { ...f }]));

            const potentialMappedFields = columnMappings
               .filter(m => m.mappedField !== null)
               .map((m, index) => {
                   const dataType = m.dataType ?? null;
                   let existingField = prevConfig.fields.find(f => !f.isStatic && !f.isCalculated && f.mappedField === m.mappedField);
                   const fieldId = existingField?.id ?? `mapped-${m.mappedField!}-${Date.now()}`;

                    let baseField: OutputFieldConfig = {
                       id: fieldId,
                       order: existingField?.order ?? (prevConfig.fields.length + index),
                       isStatic: false,
                       isCalculated: false,
                       mappedField: m.mappedField!,
                       length: prevConfig.format === 'txt' ? (existingField?.length ?? (dataType === 'Alfanumérico' ? (m.length ?? 10) : 10)) : existingField?.length,
                       paddingChar: prevConfig.format === 'txt' ? (existingField?.paddingChar ?? getDefaultPaddingChar({isStatic: false, isCalculated: false, mappedField: m.mappedField!, id: '', order: 0 }, columnMappings)) : existingField?.paddingChar,
                       paddingDirection: prevConfig.format === 'txt' ? (existingField?.paddingDirection ?? getDefaultPaddingDirection({isStatic: false, isCalculated: false, mappedField: m.mappedField!, id: '', order: 0 }, columnMappings)) : existingField?.paddingDirection,
                       dateFormat: dataType === 'Data' ? (existingField?.dateFormat ?? 'DDMMYYYY') : undefined, // Default to DDMMYYYY
                       isStaticFromPreset: existingField?.isStaticFromPreset ?? false,
                   };

                    if (prevConfig.format !== 'txt') {
                        delete baseField.length;
                        delete baseField.paddingChar;
                        delete baseField.paddingDirection;
                    }
                     if (dataType !== 'Data') {
                          delete baseField.dateFormat;
                     }

                   return baseField;
               });

           const uniqueMappedFieldsMap = new Map<string, OutputFieldConfig>();
           prevConfig.fields.forEach(f => {
               if (!f.isStatic && !f.isCalculated && f.mappedField) {
                   uniqueMappedFieldsMap.set(f.mappedField, f);
               }
           });
           potentialMappedFields.forEach(f => {
                if (!f.isStatic && !f.isCalculated && f.mappedField && !uniqueMappedFieldsMap.has(f.mappedField)) {
                   uniqueMappedFieldsMap.set(f.mappedField, f);
                } else if (!f.isStatic && !f.isCalculated && f.mappedField && uniqueMappedFieldsMap.has(f.mappedField)) {
                     const existing = uniqueMappedFieldsMap.get(f.mappedField)!;
                      const currentMapping = columnMappings.find(cm => cm.mappedField === f.mappedField);
                      const currentDataType = currentMapping?.dataType ?? null;

                      existing.dateFormat = currentDataType === 'Data' ? (existing.dateFormat ?? 'DDMMYYYY') : undefined; // Default DDMMYYYY
                       if (prevConfig.format === 'txt') {
                           existing.length = existing.length ?? (currentDataType === 'Alfanumérico' ? (currentMapping?.length ?? 10) : 10);
                           existing.paddingChar = existing.paddingChar ?? getDefaultPaddingChar(existing, columnMappings);
                           existing.paddingDirection = existing.paddingDirection ?? getDefaultPaddingDirection(existing, columnMappings);
                       } else {
                            delete existing.length;
                            delete existing.paddingChar;
                            delete existing.paddingDirection;
                       }
                        if (currentDataType !== 'Data') {
                            delete existing.dateFormat;
                       }

                      uniqueMappedFieldsMap.set(f.mappedField, existing);
                }
           });
           const uniqueMappedFields = Array.from(uniqueMappedFieldsMap.values());


            const updatedStaticFields = prevConfig.fields
                .filter((f): f is OutputFieldConfig & { isStatic: true } => f.isStatic)
                .map(f => {
                    const originalField = existingFieldsMap.get(f.id);
                    let updatedField = { ...f };

                    if (prevConfig.format === 'txt') {
                        updatedField.length = updatedField.length ?? originalField?.length ?? updatedField.staticValue?.length ?? 10;
                        updatedField.paddingChar = updatedField.paddingChar ?? originalField?.paddingChar ?? getDefaultPaddingChar(updatedField, columnMappings);
                        updatedField.paddingDirection = updatedField.paddingDirection ?? originalField?.paddingDirection ?? getDefaultPaddingDirection(updatedField, columnMappings);
                    } else {
                         delete updatedField.length;
                         delete updatedField.paddingChar;
                         delete updatedField.paddingDirection;
                    }
                    return updatedField;
                });

            const updatedCalculatedFields = prevConfig.fields
                .filter((f): f is OutputFieldConfig & { isCalculated: true } => f.isCalculated)
                .map(f => {
                    const originalField = existingFieldsMap.get(f.id);
                    let updatedField = { ...f };
                    const isDateFieldCalc = updatedField.type === 'CalculateStartDate' || updatedField.type === 'FormatPeriodMMAAAA';


                    if (prevConfig.format === 'txt') {
                         updatedField.length = updatedField.length ?? originalField?.length ?? (updatedField.type === 'CalculateStartDate' ? 8 : (updatedField.type === 'FormatPeriodMMAAAA' ? 6 : (updatedField.type === 'CalculateSituacaoRetorno' ? 1 : 10) ) );
                         updatedField.paddingChar = updatedField.paddingChar ?? originalField?.paddingChar ?? getDefaultPaddingChar(updatedField, columnMappings);
                         updatedField.paddingDirection = updatedField.paddingDirection ?? originalField?.paddingDirection ?? getDefaultPaddingDirection(updatedField, columnMappings);
                    } else {
                         delete updatedField.length;
                         delete updatedField.paddingChar;
                         delete updatedField.paddingDirection;
                    }
                    if (isDateFieldCalc) {
                        updatedField.dateFormat = updatedField.dateFormat ?? originalField?.dateFormat ?? 'DDMMYYYY'; // Default DDMMYYYY
                    } else if (updatedField.type !== 'FormatPeriodMMAAAA' && updatedField.type !== 'CalculateSituacaoRetorno') {
                         delete updatedField.dateFormat;
                    }
                    return updatedField;
                });


            let combinedFields: OutputFieldConfig[] = [
               ...updatedStaticFields,
               ...updatedCalculatedFields,
               ...uniqueMappedFields
           ];

             combinedFields = combinedFields.filter(field =>
                field.isStatic || field.isCalculated ||
                columnMappings.some(cm => cm.mappedField === field.mappedField)
            );
             combinedFields = combinedFields.filter(field =>
                !field.isCalculated ||
                field.requiredInputFields.every(reqId =>
                     columnMappings.some(cm => cm.mappedField === reqId)
                 ) || field.type === 'FormatPeriodMMAAAA' // FormatPeriodMMAAAA might not have required *mapped* inputs if using current date
            );


           combinedFields.sort((a, b) => a.order - b.order);
           const reorderedFinalFields = combinedFields.map((f, idx) => ({ ...f, order: idx }));


            const hasChanged = JSON.stringify(prevConfig.fields) !== JSON.stringify(reorderedFinalFields);

           if (hasChanged) {
               return {
                   ...prevConfig,
                   fields: reorderedFinalFields
               };
           } else {
               return prevConfig;
           }
       });
   }, [columnMappings, fileData.length, outputConfig.format, selectedLayoutPresetId]);


    const calculateFieldValue = (field: OutputFieldConfig, row: any): string => {
        if (!field.isCalculated) return '';

        try {
            switch (field.type) {
                case 'CalculateStartDate':
                    const periodStr = field.parameters?.period as string;
                    const parcelasPagasFieldId = field.requiredInputFields[0];
                    const parcelasMapping = columnMappings.find(m => m.mappedField === parcelasPagasFieldId);

                    if (!periodStr || !parcelasMapping) {
                        console.warn(`Campo Calculado ${field.fieldName}: Período ou Mapeamento de Parcelas não encontrado.`);
                        return '';
                    }
                     const parsedPeriod = parse(periodStr, 'dd/MM/yyyy', new Date());
                     if (!dateFnsIsValid(parsedPeriod)) {
                         toast({ title: "Erro de Cálculo", description: `Período '${periodStr}' inválido para o campo calculado '${field.fieldName}'. Use DD/MM/AAAA.`, variant: "destructive" });
                         return '';
                     }

                    const parcelasValueRaw = row[parcelasMapping.originalHeader];
                    let parcelasPagasNum = parseInt(removeMaskHelper(String(parcelasValueRaw ?? '0'), 'Inteiro'), 10);

                    if (isNaN(parcelasPagasNum) || parcelasPagasNum < 0) {
                        console.warn(`Campo Calculado ${field.fieldName}: Valor de Parcelas Pagas inválido ou não numérico ('${parcelasValueRaw}'). Usando 0.`);
                        parcelasPagasNum = 0;
                    }
                    const startDate = subMonths(parsedPeriod, parcelasPagasNum);
                    const dateFormatStr = field.dateFormat === 'YYYYMMDD' ? 'yyyyMMdd' : 'ddMMyyyy';
                    return format(startDate, dateFormatStr);

                case 'CalculateSituacaoRetorno':
                    const valorRealizadoFieldId = field.requiredInputFields[0];
                    const valorRealizadoMapping = columnMappings.find(m => m.mappedField === valorRealizadoFieldId);
                    if (!valorRealizadoMapping) {
                        console.warn(`Campo Calculado ${field.fieldName}: Mapeamento de Valor Realizado não encontrado.`);
                        return '';
                    }
                    const valorRealizadoRaw = row[valorRealizadoMapping.originalHeader] ?? '0';
                    let cleanedValorRealizado = String(valorRealizadoRaw).replace(/[^\d.,-]/g, '').replace(',', '.');
                    const numValorRealizado = parseFloat(cleanedValorRealizado);
                    return (!isNaN(numValorRealizado) && numValorRealizado > 0) ? 'T' : 'R';

                case 'FormatPeriodMMAAAA':
                     // Try to use mapped "Período" field first
                    const periodoFieldId = field.requiredInputFields.find(id => id === 'periodo'); // Assuming 'periodo' is the ID
                    const periodoMapping = periodoFieldId ? columnMappings.find(m => m.mappedField === periodoFieldId) : undefined;
                    let referenceDate = new Date(); // Default to current date

                    if (periodoMapping && row[periodoMapping.originalHeader]) {
                        const periodoValue = String(row[periodoMapping.originalHeader]);
                        // Attempt to parse the mapped period value flexibly
                        const commonDateFormats = ['dd/MM/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd', 'yyyyMMdd', 'ddMMyyyy'];
                        let parsedInputDate: Date | null = null;
                        for (const fmt of commonDateFormats) {
                            parsedInputDate = parse(periodoValue, fmt, new Date());
                            if (dateFnsIsValid(parsedInputDate)) {
                                referenceDate = parsedInputDate;
                                break;
                            }
                        }
                        if (!parsedInputDate || !dateFnsIsValid(parsedInputDate)) {
                             console.warn(`FormatPeriodMMAAAA: Período mapeado '${periodoValue}' inválido. Usando data atual.`);
                        }
                    } else if (field.parameters?.periodForMMAAAA) { // Fallback to parameter if provided (legacy or specific use)
                        const referenceDateStr = field.parameters.periodForMMAAAA as string;
                        const parsedRefDate = parse(referenceDateStr, 'dd/MM/yyyy', new Date());
                        if (dateFnsIsValid(parsedRefDate)) {
                            referenceDate = parsedRefDate;
                        } else {
                            console.warn(`FormatPeriodMMAAAA: Período do parâmetro '${referenceDateStr}' inválido. Usando data atual.`);
                        }
                    }
                    return format(referenceDate, 'MMyyyy');


                default:
                    console.warn(`Tipo de campo calculado desconhecido: ${field.type}`);
                    return '';
            }
        } catch (error: any) {
             console.error(`Erro ao calcular campo ${field.fieldName}:`, error);
             return '';
        }
    };

  const convertFile = () => {
    setIsProcessing(true);
    setProcessingMessage('Convertendo arquivo...');
    setConvertedData('');

     const calculatedFieldsWithIssues = outputConfig.fields
        .filter((f): f is OutputFieldConfig & { isCalculated: true } => f.isCalculated)
        .filter(cf => {
            if (cf.type !== 'FormatPeriodMMAAAA' && cf.requiredInputFields.some(reqId => !columnMappings.some(cm => cm.mappedField === reqId))) {
                return true; // Missing mapped inputs for types other than FormatPeriodMMAAAA (which can use current date)
            }
            if (cf.type === 'CalculateStartDate' && (!cf.parameters?.period || !dateFnsIsValid(parse(cf.parameters.period, 'dd/MM/yyyy', new Date())))) {
                return true;
            }
            return false;
        });

     if (calculatedFieldsWithIssues.length > 0) {
         const issueDescriptions = calculatedFieldsWithIssues.map(cf => {
             const missingInputs = cf.requiredInputFields
                 .filter(reqId => !columnMappings.some(cm => cm.mappedField === reqId))
                 .map(reqId => predefinedFields.find(pf => pf.id === reqId)?.name || reqId)
                 .join(', ');
             let description = `Campo calculado "${cf.fieldName}": `;
             if (missingInputs && cf.type !== 'FormatPeriodMMAAAA') description += `Inputs não mapeados: ${missingInputs}. `;
             if (cf.type === 'CalculateStartDate' && (!cf.parameters?.period || !dateFnsIsValid(parse(cf.parameters.period, 'dd/MM/yyyy', new Date())))) {
                 description += `Parâmetro "Período Atual" inválido ou não definido.`;
             }
             return description.trim();
         }).filter(desc => desc.endsWith('.')).join('\n'); // Only join valid descriptions

         if (issueDescriptions) {
            toast({
                title: "Erro de Configuração",
                description: (
                    <div className="flex items-start text-sm">
                        <AlertTriangle className="h-5 w-5 mr-2 flex-shrink-0 text-destructive-foreground" />
                        <span className="flex-grow">{`Problemas com campos calculados:\n${issueDescriptions}`}</span>
                    </div>
                ),
                variant: "destructive",
                duration: 15000,
            });
            setIsProcessing(false);
            setActiveTab(calculatedFieldsWithIssues.some(cf => cf.requiredInputFields.some(reqId => !columnMappings.some(cm => cm.mappedField === reqId))) ? "mapping" : "config");
            return;
         }
     }


    if (!fileData && outputConfig.fields.every(f => f.isStatic === false)) {
        toast({ title: "Erro", description: "Nenhum dado de entrada ou campo mapeado/calculado para converter.", variant: "destructive" });
        setIsProcessing(false);
        return;
    }
     if (outputConfig.fields.length === 0) {
         toast({ title: "Erro", description: "Configure os campos de saída antes de converter.", variant: "destructive" });
        setIsProcessing(false);
        setActiveTab("config");
        return;
    }

    const mappedOutputFields = outputConfig.fields.filter(f => !f.isStatic && !f.isCalculated);
    const usedMappedFields = new Set(mappedOutputFields.map(f => f.mappedField));
    const mappingsUsedInOutput = columnMappings.filter(m => m.mappedField && usedMappedFields.has(m.mappedField));

    if (mappingsUsedInOutput.some(m => !m.dataType)) {
        toast({ title: "Erro", description: "Defina o 'Tipo' para todos os campos mapeados usados na saída (Aba 2).", variant: "destructive", duration: 7000 });
        setIsProcessing(false);
        setActiveTab("mapping");
        return;
    }
     if (outputConfig.format === 'txt' && outputConfig.fields.some(f => (f.length === undefined || f.length === null || f.length <= 0) )) {
        toast({ title: "Erro", description: "Defina um 'Tamanho' válido (> 0) para todos os campos na saída TXT (Aba 3).", variant: "destructive", duration: 7000 });
        setIsProcessing(false);
        setActiveTab("config");
        return;
    }
      if (outputConfig.format === 'txt' && outputConfig.fields.some(f => !f.paddingChar || f.paddingChar.length !== 1)) {
         toast({ title: "Erro", description: "Defina um 'Caractere de Preenchimento' válido (1 caractere) para todos os campos na saída TXT (Aba 3).", variant: "destructive", duration: 7000 });
         setIsProcessing(false);
         setActiveTab("config");
         return;
     }
     if (outputConfig.format === 'csv' && (!outputConfig.delimiter || outputConfig.delimiter.length === 0)) {
        toast({ title: "Erro", description: "Defina um 'Delimitador' para a saída CSV (Aba 3).", variant: "destructive", duration: 7000 });
        setIsProcessing(false);
        setActiveTab("config");
        return;
    }
    if (outputConfig.fields.some(f => getOutputFieldDataType(f) === 'Data' && !f.dateFormat)) {
        toast({ title: "Erro", description: "Selecione um 'Formato Data' para todos os campos do tipo Data (mapeados ou calculados) na saída (Aba 3).", variant: "destructive", duration: 7000 });
        setIsProcessing(false);
        setActiveTab("config");
        return;
    }


    try {
      let resultString = '';
      const sortedOutputFields = [...outputConfig.fields].sort((a, b) => a.order - b.order);

      const dataToProcess = fileData && fileData.length > 0 ? fileData : [{}];

      dataToProcess.forEach(row => {
        let line = '';
        sortedOutputFields.forEach((outputField, fieldIndex) => {
          let value = '';
          let mapping: ColumnMapping | undefined;
          let dataType: DataType | 'Calculado' | null = null;
          let originalValue: any = null;

          if (outputField.isStatic) {
             value = outputField.staticValue ?? '';
             originalValue = value;
              if (outputConfig.format === 'txt') {
                 dataType = /^-?\d+$/.test(value) ? 'Inteiro' : /^-?\d+(\.|,)\d+$/.test(value.replace(',', '.')) ? 'Numérico' : 'Alfanumérico';
             }
          } else if (outputField.isCalculated) {
             value = calculateFieldValue(outputField, row);
             originalValue = `Calculado: ${value}`;
             dataType = getOutputFieldDataType(outputField);

             const dateFormat = outputField.dateFormat;
             if (dataType === 'Data' && value && dateFormat) {
                 try {
                      const testParseFormat = dateFormat === 'YYYYMMDD' ? 'yyyyMMdd' : 'ddMMyyyy';
                      const parsedCalcDate = parse(value, testParseFormat, new Date());
                     if (!dateFnsIsValid(parsedCalcDate)) {
                          console.warn(`Valor calculado de data '${value}' parece inválido para o formato ${dateFormat}. Gerando vazio.`);
                           value = '';
                     }
                 } catch (e) {
                     console.error(`Erro ao re-validar data calculada ${value}`, e);
                     value = '';
                 }
             }

          } else {
             mapping = columnMappings.find(m => m.mappedField === outputField.mappedField);
             if (!mapping || !mapping.originalHeader) {
                 if(fileData && fileData.length > 0 && mapping?.mappedField) {
                     console.warn(`Mapeamento ou cabeçalho original não encontrado para o campo de saída: ${outputField.mappedField}`);
                 }
                 value = '';
             } else if (!(mapping.originalHeader in row) && fileData && fileData.length > 0) {
                  console.warn(`Cabeçalho original "${mapping.originalHeader}" não encontrado na linha de dados para o campo mapeado: ${outputField.mappedField}.`);
                  value = '';
             }
              else {
                 originalValue = row[mapping.originalHeader] ?? '';
                 value = String(originalValue).trim();
                 dataType = mapping.dataType;

                  if (mapping.removeMask && dataType && value) {
                      value = removeMaskHelper(value, dataType);
                  }

                 switch (dataType) {
                      case 'CPF':
                      case 'CNPJ':
                      case 'Inteiro':
                           if (!mapping.removeMask && value) {
                                value = value.replace(/\D/g, '');
                            }
                           break;
                      case 'Numérico':
                             let numStr = value;
                             if (mapping.removeMask) {
                                 numStr = numStr.replace(/[^\d.-]/g, '');
                             } else {
                                  numStr = numStr.replace(/[R$ ]/g, '').replace(/\./g, (match, offset, fullStr) => {
                                     return offset === fullStr.lastIndexOf('.') ? '.' : '';
                                 }).replace(',', '.');
                             }

                             const parts = numStr.split('.');
                             if (parts.length > 2) {
                                 numStr = parts.slice(0, -1).join('') + '.' + parts[parts.length - 1];
                             }
                            const numMatch = numStr.match(/^(-?\d+\.?\d*)|(^-?\.\d+)/);
                            if (numMatch && numMatch[0]) {
                                let numVal = parseFloat(numMatch[0]);
                                if (isNaN(numVal)) {
                                    value = '0.00';
                                } else {
                                    value = numVal.toFixed(2);
                                }
                            } else if (value === '' || value === '0' || value === '-') {
                                value = '0.00';
                            } else {
                                console.warn(`Não foi possível analisar valor numérico: ${originalValue} (processado: ${value}). Usando 0.00`);
                                value = '0.00';
                            }
                          break;
                       case 'Data':
                            try {
                                let parsedDate: Date | null = null;
                                let cleanedValue = value;
                                if (mapping?.removeMask && value) {
                                    cleanedValue = value.replace(/[^\d]/g, '');
                                } else if (value) {
                                     cleanedValue = value.replace(/[-/.]/g, '');
                                }
                                const dateStringForParsing = String(originalValue).trim();
                                const outputDateFormat = outputField.dateFormat || 'DDMMYYYY'; // Default DDMMYYYY
                                if (/^\d{4}-\d{2}-\d{2}/.test(dateStringForParsing)) {
                                     parsedDate = parse(dateStringForParsing.substring(0, 10), 'yyyy-MM-dd', new Date());
                                }
                                if (!parsedDate || !dateFnsIsValid(parsedDate)) {
                                    const commonFormats = [
                                        'dd/MM/yyyy', 'd/M/yyyy', 'dd-MM-yyyy', 'd-M-yyyy',
                                        'MM/dd/yyyy', 'M/d/yyyy', 'MM-dd-yyyy', 'M-d-yyyy',
                                        'yyyy/MM/dd', 'yyyy/M/d', 'yyyy-MM-dd', 'yyyy-M-d',
                                        'dd/MM/yy', 'd/M/yy', 'dd-MM-yy', 'd-M-yy',
                                        'MM/dd/yy', 'M/d/yy', 'MM-dd-yy', 'M-d-yy',
                                        'yyyyMMdd', 'ddMMyyyy', 'yyMMdd', 'ddMMyy'
                                    ];
                                    for (const fmt of commonFormats) {
                                        parsedDate = parse(dateStringForParsing, fmt, new Date());
                                        if (dateFnsIsValid(parsedDate)) break;
                                    }
                                }
                                if ((!parsedDate || !dateFnsIsValid(parsedDate)) && cleanedValue && /^\d+$/.test(cleanedValue)) {
                                     if (cleanedValue.length === 8) {
                                          const fmt1 = outputDateFormat === 'YYYYMMDD' ? 'yyyyMMdd' : 'ddMMyyyy';
                                          const fmt2 = outputDateFormat === 'YYYYMMDD' ? 'ddMMyyyy' : 'yyyyMMdd';
                                          parsedDate = parse(cleanedValue, fmt1, new Date());
                                          if (!dateFnsIsValid(parsedDate)) {
                                              parsedDate = parse(cleanedValue, fmt2, new Date());
                                          }
                                     } else if (cleanedValue.length === 6) {
                                          const fmt1 = outputDateFormat === 'YYYYMMDD' ? 'yyMMdd' : 'ddMMyy';
                                          const fmt2 = outputDateFormat === 'YYYYMMDD' ? 'ddMMyy' : 'yyMMdd';
                                          parsedDate = parse(cleanedValue, fmt1, new Date());
                                          if (!dateFnsIsValid(parsedDate)) {
                                              parsedDate = parse(cleanedValue, fmt2, new Date());
                                          }
                                     }
                                }
                                if (parsedDate && dateFnsIsValid(parsedDate)) {
                                     const y = parsedDate.getFullYear();
                                     const m = String(parsedDate.getMonth() + 1).padStart(2, '0');
                                     const d = String(parsedDate.getDate()).padStart(2, '0');
                                     value = outputDateFormat === 'YYYYMMDD' ? `${y}${m}${d}` : `${d}${m}${y}`;
                                } else if (value) {
                                    console.warn(`Não foi possível analisar a data: ${originalValue} (limpo: ${cleanedValue}). Gerando vazio.`);
                                    value = '';
                                } else {
                                    value = '';
                                }
                            } catch (e) {
                                console.error(`Erro ao processar data: ${originalValue}`, e);
                                value = '';
                            }
                            break;
                      case 'Alfanumérico':
                      default:
                          break;
                 }
             }
          }


          if (outputConfig.format === 'txt') {
             const len = outputField.length ?? 0;
             const padChar = outputField.paddingChar || getDefaultPaddingChar(outputField, columnMappings);
             const padDir = outputField.paddingDirection || getDefaultPaddingDirection(outputField, columnMappings);
             let processedValue = String(value ?? '');

             let effectiveDataType = getOutputFieldDataType(outputField);
             if(outputField.isStatic) {
                 effectiveDataType = /^-?\d+(\.\d+)?$/.test(outputField.staticValue.replace(',', '.')) ? 'Numérico' : /^-?\d+$/.test(outputField.staticValue) ? 'Inteiro' : 'Alfanumérico';
             }

             if (len > 0) {
                 const isNegative = processedValue.startsWith('-');
                 let coreValue = isNegative ? processedValue.substring(1) : processedValue;
                 let sign = isNegative ? '-' : '';

                 let valueForLengthCheck = coreValue;
                 if (effectiveDataType === 'Numérico' && padChar === '0') {
                     valueForLengthCheck = coreValue.replace('.', '');
                 }


                 if (valueForLengthCheck.length > (isNegative ? len - 1 : len) ) {
                     console.warn(`Truncando valor "${processedValue}" para o campo ${outputField.isStatic ? outputField.fieldName : outputField.isCalculated ? outputField.fieldName : outputField.mappedField} pois excede o tamanho ${len}`);
                     if (padDir === 'left' && (effectiveDataType === 'Numérico' || effectiveDataType === 'Inteiro' || effectiveDataType === 'CPF' || effectiveDataType === 'CNPJ' || effectiveDataType === 'Data')) {
                         valueForLengthCheck = valueForLengthCheck.slice(-(isNegative ? len - 1 : len));
                     } else {
                         valueForLengthCheck = valueForLengthCheck.substring(0, (isNegative ? len -1 : len));
                     }
                 }

                 if (effectiveDataType === 'Numérico' && padChar === '0' && coreValue.includes('.')) {
                      const corePadLen = (isNegative ? len - 1 : len) - valueForLengthCheck.length;
                      let paddedCoreNumStr = "";
                      if (corePadLen > 0) {
                          if (padDir === 'left') {
                            paddedCoreNumStr = padChar.repeat(corePadLen) + valueForLengthCheck;
                          } else {
                            paddedCoreNumStr = valueForLengthCheck + padChar.repeat(corePadLen);
                          }
                      } else {
                        paddedCoreNumStr = valueForLengthCheck;
                      }
                      if (paddedCoreNumStr.length > 2) {
                        coreValue = paddedCoreNumStr.slice(0, -2) + "." + paddedCoreNumStr.slice(-2);
                      } else if (paddedCoreNumStr.length > 0) {
                        coreValue = "0." + paddedCoreNumStr.padStart(2, '0');
                      } else {
                        coreValue = "0.00";
                      }

                 } else {
                      const corePadLen = (isNegative ? len - 1 : len) - valueForLengthCheck.length;
                      if (corePadLen > 0) {
                          if (padDir === 'left') {
                            coreValue = padChar.repeat(corePadLen) + valueForLengthCheck;
                          } else {
                            coreValue = valueForLengthCheck + padChar.repeat(corePadLen);
                          }
                      } else {
                        coreValue = valueForLengthCheck;
                      }
                 }

                 processedValue = sign + coreValue;

                 if(processedValue.length > len){
                     console.warn(`Re-truncando valor "${processedValue}" para o tamanho ${len} após preenchimento.`);
                     if (padDir === 'left' && (effectiveDataType === 'Numérico' || effectiveDataType === 'Inteiro' || effectiveDataType === 'CPF' || effectiveDataType === 'CNPJ' || effectiveDataType === 'Data')) {
                          processedValue = processedValue.slice(-len);
                      } else {
                           processedValue = processedValue.slice(0, len);
                      }
                 }
             } else {
                 processedValue = '';
             }
             line += processedValue;

          } else if (outputConfig.format === 'csv') {
            if (fieldIndex > 0) {
              line += outputConfig.delimiter;
            }
             let csvValue = String(value ?? '');
             const effectiveDataType = getOutputFieldDataType(outputField);
              if (effectiveDataType === 'Numérico') {
                    csvValue = csvValue.replace('.', ',');
              }
             const needsQuotes = csvValue.includes(outputConfig.delimiter!) || csvValue.includes('"') || csvValue.includes('\n');
             if (needsQuotes) {
                csvValue = `"${csvValue.replace(/"/g, '""')}"`;
            }
            line += csvValue;
          }
        });
        resultString += line + '\n';
      });

        const resultBuffer = iconv.encode(resultString.trimEnd(), outputConfig.encoding);
        setConvertedData(resultBuffer);

      setActiveTab("result");
      toast({ title: "Sucesso", description: "Arquivo convertido com sucesso!" });
    } catch (error: any) {
      console.error("Erro na conversão:", error);
      toast({
        title: "Erro na Conversão",
        description: error.message || "Ocorreu um erro inesperado durante a conversão.",
        variant: "destructive",
      });
       setActiveTab("config");
    } finally {
      setIsProcessing(false);
      setProcessingMessage('Processando...');
    }
  };

   const openDownloadDialog = () => {
        if (!convertedData) return;
        const baseFileName = fileName ? fileName.split('.').slice(0, -1).join('.') : 'arquivo';
        const proposed = `${baseFileName}_convertido.${outputConfig.format}`;
        setDownloadDialogState({
            isOpen: true,
            proposedFilename: proposed,
            finalFilename: proposed,
        });
    };

    const handleDownloadFilenameChange = (event: React.ChangeEvent<HTMLInputElement>) => {
         setDownloadDialogState(prev => ({
             ...prev,
             finalFilename: event.target.value,
         }));
     };


   const confirmDownload = () => {
        const { finalFilename } = downloadDialogState;
        if (!convertedData || !finalFilename.trim()) {
            toast({ title: "Erro", description: "Nome do arquivo não pode ser vazio.", variant: "destructive" });
            return;
        }

        const cleanFilename = finalFilename.trim();
        const finalFilenameWithExt = cleanFilename.endsWith(`.${outputConfig.format}`)
            ? cleanFilename
            : `${cleanFilename}.${outputConfig.format}`;


        const encoding = outputConfig.encoding.toLowerCase();
        const mimeType = outputConfig.format === 'txt'
            ? `text/plain;charset=${encoding}`
            : `text/csv;charset=${encoding}`;

         const blob = convertedData instanceof Buffer
             ? new Blob([convertedData], { type: mimeType })
             : new Blob([String(convertedData)], { type: mimeType });

        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = finalFilenameWithExt;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        toast({ title: "Download Iniciado", description: `Arquivo ${link.download} sendo baixado.`});
        setDownloadDialogState({ isOpen: false, proposedFilename: '', finalFilename: '' });
    };

    const openConfigDialog = (action: 'save' | 'load') => {
        setConfigManagementDialogState({
            isOpen: true,
            action: action,
            configName: action === 'save' ? (outputConfig.name || `Nova Config ${savedConfigs.length + 1}`) : '',
            selectedConfigToLoad: null,
        });
    };

    const handleConfigDialogChange = (field: keyof ConfigManagementDialogState, value: any) => {
        setConfigManagementDialogState(prev => ({
            ...prev,
            [field]: value === NONE_VALUE_PLACEHOLDER ? null : value,
        }));
    };

    const saveCurrentConfig = () => {
        const { configName } = configManagementDialogState;
        if (!configName.trim()) {
            toast({ title: "Erro", description: "Nome da configuração não pode ser vazio.", variant: "destructive" });
            return;
        }

        const configToSave: OutputConfig = {
            ...outputConfig,
            name: configName.trim(),
        };

        setSavedConfigs(prev => {
            const existingIndex = prev.findIndex(c => c.name === configToSave.name);
            let updatedConfigs;
            if (existingIndex > -1) {
                updatedConfigs = [...prev];
                updatedConfigs[existingIndex] = configToSave;
                 toast({ title: "Sucesso", description: `Configuração "${configToSave.name}" atualizada.` });
            } else {
                updatedConfigs = [...prev, configToSave];
                 toast({ title: "Sucesso", description: `Configuração "${configToSave.name}" salva.` });
            }
            saveAllConfigs(updatedConfigs);
            return updatedConfigs;
        });
        setOutputConfig(configToSave);
        setConfigManagementDialogState({ isOpen: false, action: null, configName: '', selectedConfigToLoad: null });
    };

    const loadSelectedConfig = () => {
        const { selectedConfigToLoad } = configManagementDialogState;
        if (!selectedConfigToLoad) {
            toast({ title: "Erro", description: "Selecione uma configuração para carregar.", variant: "destructive" });
            return;
        }

        const config = savedConfigs.find(c => c.name === selectedConfigToLoad);
        if (!config) {
            toast({ title: "Erro", description: "Configuração selecionada não encontrada.", variant: "destructive" });
            return;
        }

        setOutputConfig(config);
        setSelectedLayoutPresetId("custom");
        setConfigManagementDialogState({ isOpen: false, action: null, configName: '', selectedConfigToLoad: null });
        toast({ title: "Sucesso", description: `Configuração "${config.name}" carregada.` });
    };

    const deleteConfig = (configNameToDelete: string) => {
        if (!configNameToDelete) return;
         const confirmed = window.confirm(`Tem certeza que deseja excluir a configuração "${configNameToDelete}"? Esta ação não pode ser desfeita.`);
        if (!confirmed) return;

        setSavedConfigs(prev => {
            const updatedConfigs = prev.filter(c => c.name !== configNameToDelete);
            saveAllConfigs(updatedConfigs);
            return updatedConfigs;
        });
         setConfigManagementDialogState(prev => ({
             ...prev,
             selectedConfigToLoad: prev.selectedConfigToLoad === configNameToDelete ? null : prev.selectedConfigToLoad
         }));
        toast({ title: "Sucesso", description: `Configuração "${configNameToDelete}" excluída.` });
    };

    const applyLayoutPreset = useCallback((presetId: string) => {
        if (presetId === "custom") {
            return;
        }
        const preset = LAYOUT_PRESETS.find(p => p.id === presetId);
        if (!preset) {
            toast({ title: "Erro", description: "Padrão de Layout não encontrado.", variant: "destructive" });
            return;
        }

        const newOutputFields: OutputFieldConfig[] = preset.fields.map((presetField, index) => {
            const baseId = `preset-${preset.id}-${presetField.targetMappedId || presetField.staticFieldName?.replace(/\s+/g, '') || presetField.calculatedInternalName}-${index}`;
            let isStaticFromPresetFlag = false;

            if (presetField.targetMappedId) {
                const isMapped = columnMappings.some(cm => cm.mappedField === presetField.targetMappedId);
                // Handle default static values for specific unmapped fields within presets
                if (presetField.targetMappedId === 'estabelecimento_empresa' && !isMapped) {
                    isStaticFromPresetFlag = true;
                    return {
                        id: `${baseId}-static-estab`, isStatic: true, isCalculated: false, fieldName: 'Estabelecimento/Empresa (Padrão)', staticValue: '001',
                        order: index, length: presetField.length, paddingChar: '0', paddingDirection: 'left', isStaticFromPreset: isStaticFromPresetFlag
                    };
                }
                if (presetField.targetMappedId === 'orgao_filial' && !isMapped) {
                     isStaticFromPresetFlag = true;
                    return {
                        id: `${baseId}-static-orgao`, isStatic: true, isCalculated: false, fieldName: 'Órgão/Filial (Padrão)', staticValue: '001',
                        order: index, length: presetField.length, paddingChar: '0', paddingDirection: 'left', isStaticFromPreset: isStaticFromPresetFlag
                    };
                }
                 if (presetField.targetMappedId === 'data_fim_contrato' && !isMapped && (preset.id === "margem_simples_econsig" || preset.id === "margem_cartao_econsig")) {
                     isStaticFromPresetFlag = true;
                     return {
                         id: `${baseId}-static-dtfim`, isStatic: true, isCalculated: false, fieldName: 'Data Fim Contrato (Padrão)', staticValue: '00000000',
                         order: index, length: presetField.length, paddingChar: '0', paddingDirection: 'left', dateFormat: presetField.dateFormat, isStaticFromPreset: isStaticFromPresetFlag
                     };
                 }
                return {
                    id: baseId, isStatic: false, isCalculated: false, mappedField: presetField.targetMappedId,
                    order: index, length: presetField.length, paddingChar: presetField.paddingChar, paddingDirection: presetField.paddingDirection,
                    dateFormat: presetField.dateFormat, isStaticFromPreset: false
                };
            } else if (presetField.staticFieldName && presetField.staticValue !== undefined) {
                 isStaticFromPresetFlag = true;
                return {
                    id: baseId, isStatic: true, isCalculated: false, fieldName: presetField.staticFieldName, staticValue: presetField.staticValue,
                    order: index, length: presetField.length, paddingChar: presetField.paddingChar, paddingDirection: presetField.paddingDirection,
                    dateFormat: presetField.dateFormat, isStaticFromPreset: isStaticFromPresetFlag
                };
            } else if (presetField.calculatedInternalName && presetField.calculationType && presetField.calculatedDisplayName) {
                 const params = presetField.calculationType === 'CalculateStartDate' ? { period: '' } : {}; // period will be asked in dialog
                return {
                    id: baseId, isStatic: false, isCalculated: true, type: presetField.calculationType, fieldName: presetField.calculatedDisplayName,
                    requiredInputFields: presetField.requiredInputFieldsForCalc || [],
                    parameters: params,
                    order: index, length: presetField.length, paddingChar: presetField.paddingChar, paddingDirection: presetField.paddingDirection,
                    dateFormat: presetField.dateFormat, isStaticFromPreset: false
                };
            }
            console.error("applyLayoutPreset: Invalid field definition in preset", presetField);
            return { id: `error-preset-${index}`, isStatic: true, fieldName: 'ERRO NO PRESET', staticValue: 'X', order: index, length: 1, paddingChar: 'X', paddingDirection: 'right', isStaticFromPreset: true } as OutputFieldConfig;
        });

        setOutputConfig(prev => ({
            ...prev,
            name: preset.name, // Set the config name to the preset name
            fields: newOutputFields
        }));

         toast({
            title: "Padrão de Layout Aplicado",
            description: (
                <div className="flex items-start text-sm">
                    <AlertTriangle className="h-5 w-5 mr-2 flex-shrink-0 text-destructive-foreground" />
                    <span className="flex-grow">O padrão "{preset.name}" foi carregado. <strong>Revise os campos de saída</strong>, pois alguns mapeamentos podem precisar de ajuste manual, especialmente se as colunas do seu arquivo forem diferentes das esperadas pelo padrão ou campos calculados exigirem parâmetros.</span>
                </div>
            ),
            variant: "destructive",
            duration: 15000,
        });


    }, [columnMappings, toast]);


  const memoizedPredefinedFields = useMemo(() => {
      const groupedFields = predefinedFields.reduce((acc, field) => {
          const group = field.group || 'Personalizado';
          if (!acc[group]) {
              acc[group] = [];
          }
          acc[group].push(field);
          return acc;
      }, {} as Record<string, PredefinedField[]>);

      for (const group in groupedFields) {
          groupedFields[group].sort((a, b) => a.name.localeCompare(b.name));
      }

      const groupOrder = ['Padrão', 'Margem', 'Histórico/Retorno', 'Principal Personalizado', 'Opcional Personalizado', 'Personalizado'];
      const sortedGroups: { groupName: string, fields: PredefinedField[] }[] = [];

      groupOrder.forEach(groupName => {
          if (groupedFields[groupName]) {
              sortedGroups.push({ groupName, fields: groupedFields[groupName] });
              delete groupedFields[groupName];
          }
      });

      Object.keys(groupedFields).sort().forEach(groupName => {
          sortedGroups.push({ groupName, fields: groupedFields[groupName] });
      });


      return sortedGroups;
  }, [predefinedFields]);


 const renderMappedOutputFieldSelect = (currentField: OutputFieldConfig & { isStatic: false, isCalculated: false }) => {
     const currentFieldMappedId = currentField.mappedField;

     const availableOptions = memoizedPredefinedFields.flatMap(group =>
        group.fields
            .filter(pf =>
                columnMappings.some(cm => cm.mappedField === pf.id)
            )
            .filter(pf =>
                 pf.id === currentFieldMappedId ||
                 !outputConfig.fields.some(of => !of.isStatic && !of.isCalculated && of.mappedField === pf.id)
             )
     );


     return (
         <Select
             value={currentFieldMappedId || NONE_VALUE_PLACEHOLDER}
             onValueChange={(value) => handleOutputFieldChange(currentField.id, 'mappedField', value)}
             disabled={isProcessing}
         >
             <SelectTrigger className="w-full text-xs h-8">
                 <SelectValue placeholder="Selecione o Campo" />
             </SelectTrigger>
             <SelectContent>
                 <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>-- Selecione --</SelectItem>
                 {availableOptions.length > 0 ? (
                     memoizedPredefinedFields.map(group => (
                         <SelectGroup key={group.groupName}>
                             <SelectLabel>{group.groupName}</SelectLabel>
                             {group.fields
                                 .filter(field =>
                                      field.id === currentFieldMappedId ||
                                      !outputConfig.fields.some(of => !of.isStatic && !of.isCalculated && of.mappedField === field.id)
                                  )
                                 .filter(field => columnMappings.some(cm => cm.mappedField === field.id))
                                 .map(field => (
                                     <SelectItem key={field.id} value={field.id}>
                                         {field.name}
                                     </SelectItem>
                                 ))
                             }
                         </SelectGroup>
                     ))
                 ) : (
                      <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>Nenhum campo mapeado disponível</SelectItem>
                 )}
             </SelectContent>
         </Select>
     );
 };

   const getOutputFieldDataType = (field: OutputFieldConfig): DataType | 'Calculado' | null => {
       if (field.isStatic) return null;
       if (field.isCalculated) {
            if (field.type === 'CalculateStartDate') return 'Data';
            if (field.type === 'FormatPeriodMMAAAA') return 'Data'; // Although output is MMyyyy, base is date
            if (field.type === 'CalculateSituacaoRetorno') return 'Alfanumérico';
            return 'Calculado';
       }
       const mapping = columnMappings.find(cm => cm.mappedField === field.mappedField);
       return mapping?.dataType ?? null;
   };

  return (
    <div className="container mx-auto p-4 md:p-8 flex flex-col items-center min-h-screen bg-background">
      <Card className="w-full max-w-5xl shadow-lg">
        <CardHeader className="text-center">
          <CardTitle className="text-3xl font-bold text-foreground">
            <Columns className="inline-block mr-2 text-accent" /> SCA - Sistema para conversão de arquivos
          </CardTitle>
          <CardDescription className="text-muted-foreground">
            Converta seus arquivos Excel (XLS, XLSX, ODS), CSV ou CNAB240 (.RET) para layouts TXT ou CSV personalizados.
            <br />
             <span className="text-xs italic flex items-center justify-center mt-1">
               <Info className="w-3 h-3 mr-1" /> Seus dados não são armazenados, garantindo conformidade com a LGPD.
             </span>
          </CardDescription>
        </CardHeader>

        <CardContent>
          <Tabs value={activeTab} onValueChange={setActiveTab} className="w-full">
             <TabsList className="grid w-full grid-cols-4 mb-6">
                 <TabsTrigger value="upload" disabled={isProcessing} data-state={activeTab === 'upload' ? 'active' : 'inactive'} className={activeTab === 'upload' ? 'tabs-trigger-active' : ''}>1. Upload</TabsTrigger>
                 <TabsTrigger value="mapping" disabled={isProcessing || !file} data-state={activeTab === 'mapping' ? 'active' : 'inactive'} className={activeTab === 'mapping' ? 'tabs-trigger-active' : ''}>2. Mapeamento</TabsTrigger>
                 <TabsTrigger value="config" disabled={isProcessing || !file } data-state={activeTab === 'config' ? 'active' : 'inactive'} className={activeTab === 'config' ? 'tabs-trigger-active' : ''}>3. Configurar Saída</TabsTrigger>
                 <TabsTrigger value="result" disabled={isProcessing || !convertedData} data-state={activeTab === 'result' ? 'active' : 'inactive'} className={activeTab === 'result' ? 'tabs-trigger-active' : ''}>4. Resultado</TabsTrigger>
             </TabsList>

            {/* 1. Upload Tab */}
            <TabsContent value="upload">
              <div className="flex flex-col items-center space-y-6 p-6 border rounded-lg bg-card">
                  <Label htmlFor="file-upload" className="text-lg font-semibold text-foreground cursor-pointer hover:text-accent transition-colors">
                     <Button asChild variant="default" className="bg-accent hover:bg-accent/90 text-accent-foreground cursor-pointer">
                          <span>
                              <Upload className="mr-2 h-5 w-5 inline-block" />
                               Selecione o Arquivo para Conversão
                          </span>
                     </Button>
                     <Input
                        id="file-upload"
                        type="file"
                        accept=".xls,.xlsx,.ods,.csv,.ret"
                        onChange={handleFileChange}
                        className="hidden"
                        disabled={isProcessing}
                     />
                 </Label>

                <p className="text-sm text-muted-foreground">Formatos suportados: XLS, XLSX, ODS, CSV, RET (CNAB240 - básico)</p>

                {fileName && (
                  <div className="mt-4 text-center text-sm text-muted-foreground">
                    Arquivo selecionado: <span className="font-medium text-foreground">{fileName}</span>
                  </div>
                )}
                 {isProcessing && activeTab === "upload" && (
                    <div className="flex items-center text-accent animate-pulse">
                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                        {processingMessage}
                    </div>
                  )}
              </div>
            </TabsContent>

            {/* 2. Mapping Tab */}
            <TabsContent value="mapping">
              {isProcessing && activeTab === "mapping" && (
                 <div className="flex items-center justify-center text-accent animate-pulse p-4">
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      {processingMessage}
                  </div>
               )}
              {!isProcessing && headers.length > 0 && (
                <div className="space-y-6">
                  <Card>
                     <CardHeader>
                         <CardTitle className="text-xl">Mapeamento de Colunas de Entrada</CardTitle>
                         <CardDescription>
                            Associe as colunas do seu arquivo ({headers.length} colunas detectadas | {fileData.length} linhas detectadas),
                            configure tipos, tamanhos e remoção de máscaras.
                         </CardDescription>
                     </CardHeader>
                     <CardContent>
                       <div className="flex justify-end items-center mb-4 gap-2">
                         <Label htmlFor="show-preview" className="text-sm font-medium">Mostrar Pré-visualização (5 linhas)</Label>
                         <Switch id="show-preview" checked={showPreview} onCheckedChange={setShowPreview} />
                       </div>
                       {showPreview && (
                          <div className="mb-6 max-h-60 overflow-auto border rounded-md bg-secondary/30">
                             <Table>
                                 <TableHeader>
                                     <TableRow>
                                         {headers.map((header, idx) => <TableHead key={`prev-h-${idx}`}>{header}</TableHead>)}
                                     </TableRow>
                                 </TableHeader>
                                 <TableBody>
                                     {getSampleData().map((row, rowIndex) => (
                                         <TableRow key={`prev-r-${rowIndex}`}>
                                             {headers.map((header, colIndex) => (
                                                 <TableCell key={`prev-c-${rowIndex}-${colIndex}`} className="text-xs whitespace-nowrap">
                                                    {String(row[header] ?? '').substring(0, 50)}
                                                    {String(row[header] ?? '').length > 50 ? '...' : ''}
                                                 </TableCell>
                                             ))}
                                         </TableRow>
                                     ))}
                                      {getSampleData().length === 0 && (
                                          <TableRow><TableCell colSpan={headers.length} className="text-center text-muted-foreground">Nenhuma linha de dados na pré-visualização.</TableCell></TableRow>
                                       )}
                                 </TableBody>
                             </Table>
                          </div>
                        )}

                        <div className="max-h-[45vh] overflow-auto">
                           <Table>
                             <TableHeader>
                               <TableRow>
                                 <TableHead className="w-[22%]">Coluna Original</TableHead>
                                 <TableHead className="w-[22%]">Mapear para Campo</TableHead>
                                 <TableHead className="w-[18%]">Tipo</TableHead>
                                 <TableHead className="w-[10%]">
                                     Tam.
                                    <TooltipProvider>
                                        <Tooltip>
                                            <TooltipTrigger asChild>
                                                 <Button variant="ghost" size="icon" className="ml-1 h-6 w-6 text-muted-foreground cursor-help align-middle"><HelpCircle className="h-4 w-4" /></Button>
                                            </TooltipTrigger>
                                            <TooltipContent>
                                                <p>Tamanho máx. para campos Alfanuméricos.</p>
                                                <p>Usado para definir o tamanho padrão na saída TXT.</p>
                                                <p>(Ignorado para outros tipos).</p>
                                            </TooltipContent>
                                        </Tooltip>
                                    </TooltipProvider>
                                 </TableHead>
                                 <TableHead className="w-[20%] text-center">
                                     Remover Máscara
                                      <TooltipProvider>
                                        <Tooltip>
                                            <TooltipTrigger asChild>
                                                 <Button variant="ghost" size="icon" className="ml-1 h-6 w-6 text-muted-foreground cursor-help align-middle"><HelpCircle className="h-4 w-4" /></Button>
                                            </TooltipTrigger>
                                            <TooltipContent>
                                                <p>Remove caracteres não numéricos/separadores</p>
                                                <p>antes do processamento para certos tipos.</p>
                                                <p>(Ex: CPF, CNPJ, Data, Numérico, Inteiro, RG)</p>
                                                 <p>(Padrão: Ativado para esses tipos, exceto 'Margem Bruta' e campos de Valor se Numérico)</p>
                                            </TooltipContent>
                                        </Tooltip>
                                    </TooltipProvider>
                                  </TableHead>
                               </TableRow>
                             </TableHeader>
                             <TableBody>
                               {columnMappings.map((mapping, index) => {
                                 const mappedFieldDetails = mapping.mappedField ? predefinedFields.find(pf => pf.id === mapping.mappedField) : null;
                                 const isValorOrMargemFieldNumerico = (mappedFieldDetails?.id === 'margem_bruta' || valorFieldIds.includes(mappedFieldDetails?.id || '')) && mapping.dataType === 'Numérico';
                                 return (
                                     <TableRow key={index}>
                                       <TableCell className="font-medium text-xs">{mapping.originalHeader}</TableCell>
                                       <TableCell>
                                            <div className="flex items-center gap-1">
                                                 <Select
                                                   value={mapping.mappedField || NONE_VALUE_PLACEHOLDER}
                                                   onValueChange={(value) => handleMappingChange(index, 'mappedField', value)}
                                                    disabled={isProcessing}
                                                 >
                                                   <SelectTrigger className="text-xs h-8 flex-grow">
                                                     <SelectValue placeholder="Selecione ou deixe em branco" />
                                                   </SelectTrigger>
                                                   <SelectContent>
                                                     <SelectItem value={NONE_VALUE_PLACEHOLDER}>-- Sem mapeamento --</SelectItem>
                                                      {memoizedPredefinedFields.map(group => (
                                                          <SelectGroup key={group.groupName}>
                                                              <SelectLabel>{group.groupName}</SelectLabel>
                                                              {group.fields.map(field => (
                                                                  <SelectItem key={field.id} value={field.id}>{field.name}</SelectItem>
                                                              ))}
                                                          </SelectGroup>
                                                      ))}
                                                   </SelectContent>
                                                 </Select>
                                                 {mappedFieldDetails?.comment && (
                                                      <TooltipProvider>
                                                          <Tooltip>
                                                              <TooltipTrigger asChild>
                                                                  <HelpCircle className="h-4 w-4 text-muted-foreground flex-shrink-0 cursor-help" />
                                                              </TooltipTrigger>
                                                              <TooltipContent>
                                                                  <p>{mappedFieldDetails.comment}</p>
                                                              </TooltipContent>
                                                          </Tooltip>
                                                      </TooltipProvider>
                                                  )}
                                           </div>
                                       </TableCell>
                                       <TableCell>
                                         <Select
                                           value={mapping.dataType || NONE_VALUE_PLACEHOLDER}
                                           onValueChange={(value) => handleMappingChange(index, 'dataType', value)}
                                           disabled={isProcessing || !mapping.mappedField}
                                         >
                                           <SelectTrigger className="text-xs h-8">
                                             <SelectValue placeholder="Tipo (Obrigatório se mapeado)" />
                                           </SelectTrigger>
                                           <SelectContent>
                                             <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>-- Selecione --</SelectItem>
                                             {DATA_TYPES.map(type => (
                                               <SelectItem key={type} value={type}>{type}</SelectItem>
                                             ))}
                                           </SelectContent>
                                         </Select>
                                       </TableCell>
                                       <TableCell>
                                         <Input
                                           type="number"
                                           min="1"
                                           value={mapping.length ?? ''}
                                           onChange={(e) => handleMappingChange(index, 'length', e.target.value)}
                                           placeholder="Tam."
                                           className="w-full text-xs h-8"
                                           disabled={isProcessing || !mapping.dataType || mapping.dataType !== 'Alfanumérico'}
                                         />
                                       </TableCell>
                                        <TableCell className="text-center">
                                          <Switch
                                              checked={mapping.removeMask}
                                              onCheckedChange={(checked) => handleMappingChange(index, 'removeMask', checked)}
                                              disabled={isProcessing || !mapping.dataType || mapping.dataType === 'Alfanumérico' || isValorOrMargemFieldNumerico}
                                              aria-label={`Remover máscara para ${mapping.originalHeader}`}
                                              className="scale-75"
                                          />
                                       </TableCell>
                                     </TableRow>
                                   );
                               })}
                             </TableBody>
                           </Table>
                         </div>
                     </CardContent>
                  </Card>

                  <Card>
                     <CardHeader>
                         <CardTitle className="text-xl">Gerenciar Campos Pré-definidos</CardTitle>
                         <CardDescription>
                            Adicione, edite ou remova campos para o mapeamento. Campos Principais são fixos.
                            Campos personalizados podem ser marcados como "Principais" para serem salvos no navegador para futuras conversões,
                            ou "Opcionais" para uso apenas nesta sessão (descartados ao atualizar a página se não salvos).
                         </CardDescription>
                     </CardHeader>
                      <CardContent>
                         <div className="flex justify-end mb-4">
                             <Button onClick={openAddPredefinedFieldDialog} disabled={isProcessing} variant="outline">
                                 <Plus className="mr-2 h-4 w-4" /> Adicionar Novo Campo
                             </Button>
                         </div>
                         <div className="space-y-2 max-h-40 overflow-y-auto border rounded p-2 bg-secondary/30">
                              {memoizedPredefinedFields.map(group => (
                                <React.Fragment key={group.groupName}>
                                  <h4 className="text-sm font-semibold text-muted-foreground px-2 pt-2">{group.groupName}</h4>
                                  {group.fields.map(field => (
                                     <div key={field.id} className="flex items-center justify-between p-2 border-b last:border-b-0 gap-2">
                                         <div className="flex items-center gap-1 flex-wrap flex-grow">
                                             <span className="text-sm font-medium">{field.name}</span>
                                             <span className="text-xs text-muted-foreground">({field.id})</span>
                                              <span className={`ml-1 text-xs ${field.isPersistent ? 'text-green-600 font-semibold' : 'text-yellow-600'}`}>
                                                 ({field.isPersistent ? 'Principal' : 'Opcional'})
                                               </span>
                                              {field.comment && (
                                                   <TooltipProvider>
                                                      <Tooltip>
                                                            <TooltipTrigger asChild>
                                                                <HelpCircle className="h-3 w-3 inline-block ml-1 text-muted-foreground cursor-help" />
                                                            </TooltipTrigger>
                                                            <TooltipContent><p>{field.comment}</p></TooltipContent>
                                                        </Tooltip>
                                                    </TooltipProvider>
                                               )}
                                         </div>
                                         <div className="flex gap-1 flex-shrink-0">
                                               <TooltipProvider>
                                                   <Tooltip>
                                                       <TooltipTrigger asChild>
                                                            <Button
                                                                variant="ghost"
                                                                size="icon"
                                                                onClick={() => openEditPredefinedFieldDialog(field)}
                                                                disabled={isProcessing}
                                                                className="h-7 w-7 text-muted-foreground hover:text-accent"
                                                                aria-label={`Editar campo ${field.name}`}
                                                            >
                                                                <Edit className="h-4 w-4" />
                                                            </Button>
                                                       </TooltipTrigger>
                                                       <TooltipContent><p>Editar "{field.name}"</p></TooltipContent>
                                                   </Tooltip>
                                               </TooltipProvider>
                                               <TooltipProvider>
                                                    <Tooltip>
                                                        <TooltipTrigger asChild>
                                                             <Button
                                                                 variant="ghost"
                                                                 size="icon"
                                                                 onClick={() => removePredefinedField(field.id)}
                                                                 disabled={isProcessing || field.isCore}
                                                                 className="h-7 w-7 text-muted-foreground hover:text-destructive disabled:text-muted-foreground/50 disabled:cursor-not-allowed"
                                                                 aria-label={`Remover campo ${field.name}`}
                                                             >
                                                                 <Trash2 className="h-4 w-4" />
                                                             </Button>
                                                        </TooltipTrigger>
                                                       <TooltipContent>
                                                            {field.isCore
                                                               ? <p>Não é possível remover campos principais originais.</p>
                                                               : <p>Remover campo "{field.name}"</p>
                                                            }
                                                       </TooltipContent>
                                                    </Tooltip>
                                                </TooltipProvider>
                                           </div>
                                     </div>
                                  ))}
                                </React.Fragment>
                              ))}
                             {predefinedFields.length === 0 && <p className="text-sm text-muted-foreground text-center p-2">Nenhum campo pré-definido encontrado.</p>}
                         </div>
                      </CardContent>
                       <CardFooter className="flex justify-end">
                          <Button onClick={() => setActiveTab("config")} disabled={isProcessing || columnMappings.length === 0} variant="default">
                              Próximo: Configurar Saída <ArrowRight className="ml-2 h-4 w-4" />
                          </Button>
                       </CardFooter>
                  </Card>
                </div>
              )}
               {!isProcessing && headers.length === 0 && file && (
                 <p className="text-center text-muted-foreground p-4">Nenhum cabeçalho encontrado ou arquivo ainda não processado/inválido.</p>
               )}
               {!isProcessing && !file && (
                   <p className="text-center text-muted-foreground p-4">Faça o upload de um arquivo na aba "Upload" para começar.</p>
               )}
            </TabsContent>

            {/* 3. Configuration Tab */}
            <TabsContent value="config">
              {isProcessing && activeTab === "config" && (
                 <div className="flex items-center justify-center text-accent animate-pulse p-4">
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      {processingMessage}
                  </div>
               )}
               {!isProcessing && file && (
                 <div className="space-y-6">
                    <Card>
                        <CardHeader className='pb-2'>
                             <div className="flex justify-between items-center">
                                 <div>
                                     <CardTitle className="text-xl">Configuração do Arquivo de Saída</CardTitle>
                                     <CardDescription>
                                         Modelo: <span className="font-semibold text-foreground">{selectedLayoutPresetId !== "custom" ? LAYOUT_PRESETS.find(p => p.id === selectedLayoutPresetId)?.name : (outputConfig.name || 'Personalizado')}</span>.
                                          Defina formato, codificação, delimitador (CSV), ordem e formatação dos campos.
                                     </CardDescription>
                                 </div>
                                 <div className="flex gap-2">
                                      <Button variant="outline" size="sm" onClick={() => openConfigDialog('save')} disabled={isProcessing}>
                                          <Save className="mr-2 h-4 w-4" /> Salvar Modelo
                                      </Button>
                                      <Button variant="outline" size="sm" onClick={() => openConfigDialog('load')} disabled={isProcessing || savedConfigs.length === 0}>
                                          <Server className="mr-2 h-4 w-4" /> Carregar Modelo
                                      </Button>
                                  </div>
                             </div>
                         </CardHeader>
                         <CardContent className="space-y-4 pt-4">
                             <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-end">
                                <div className="flex-1">
                                    <Label htmlFor="output-format">Formato de Saída</Label>
                                    <Select
                                        value={outputConfig.format}
                                        onValueChange={(value) => handleOutputFormatChange(value as OutputFormat)}
                                        disabled={isProcessing}
                                    >
                                        <SelectTrigger id="output-format" className="w-full">
                                            <SelectValue />
                                        </SelectTrigger>
                                        <SelectContent>
                                            <SelectItem value="txt">TXT Posicional (Largura Fixa)</SelectItem>
                                            <SelectItem value="csv">CSV (Delimitado)</SelectItem>
                                        </SelectContent>
                                    </Select>
                                </div>
                                 <div className="flex-1">
                                      <Label htmlFor="layout-preset">Padrão Layout (TXT)</Label>
                                      <Select
                                          value={selectedLayoutPresetId}
                                          onValueChange={(value) => {
                                              setSelectedLayoutPresetId(value);
                                              if (value !== "custom") {
                                                  applyLayoutPreset(value);
                                              }
                                          }}
                                          disabled={isProcessing || outputConfig.format !== 'txt'}
                                      >
                                          <SelectTrigger id="layout-preset">
                                              <SelectValue placeholder="Selecione um padrão..." />
                                          </SelectTrigger>
                                          <SelectContent>
                                              <SelectItem value="custom">Personalizado</SelectItem>
                                              {LAYOUT_PRESETS.map(preset => (
                                                  <SelectItem key={preset.id} value={preset.id}>{preset.name}</SelectItem>
                                              ))}
                                          </SelectContent>
                                      </Select>
                                  </div>
                                 <div className="flex-1">
                                      <div className="flex items-center">
                                        <Label htmlFor="output-encoding">Codificação</Label>
                                         <TooltipProvider>
                                                <Tooltip>
                                                    <TooltipTrigger asChild>
                                                         <Button variant="ghost" size="icon" className="ml-1 h-6 w-6 text-muted-foreground cursor-help align-middle"><HelpCircle className="h-4 w-4" /></Button>
                                                    </TooltipTrigger>
                                                    <TooltipContent>
                                                        <p>Define a codificação de caracteres do arquivo de saída.</p>
                                                        <p>UTF-8 é recomendado, ISO-8859-1 (Latin-1) ou Windows-1252 podem ser necessários para sistemas legados.</p>
                                                    </TooltipContent>
                                                </Tooltip>
                                        </TooltipProvider>
                                       </div>
                                    <Select
                                        value={outputConfig.encoding}
                                        onValueChange={(value) => setOutputConfig(prev => ({ ...prev, encoding: value as OutputEncoding }))}
                                        disabled={isProcessing}
                                    >
                                        <SelectTrigger id="output-encoding" className="w-full">
                                            <SelectValue />
                                        </SelectTrigger>
                                        <SelectContent>
                                            {OUTPUT_ENCODINGS.map(enc => (
                                                <SelectItem key={enc} value={enc}>{enc}</SelectItem>
                                            ))}
                                        </SelectContent>
                                    </Select>
                                </div>


                                {outputConfig.format === 'csv' && (
                                    <div className="flex-1 md:max-w-[150px] md:col-start-3">
                                        <div className="flex items-center">
                                            <Label htmlFor="csv-delimiter">Delimitador CSV</Label>
                                            <TooltipProvider>
                                                <Tooltip>
                                                    <TooltipTrigger asChild>
                                                         <Button variant="ghost" size="icon" className="ml-1 h-6 w-6 text-muted-foreground cursor-help align-middle"><HelpCircle className="h-4 w-4" /></Button>
                                                    </TooltipTrigger>
                                                    <TooltipContent>
                                                        <p>Caractere(s) para separar os campos (ex: | ; , ).</p>
                                                    </TooltipContent>
                                                </Tooltip>
                                            </TooltipProvider>
                                         </div>
                                        <Input
                                            id="csv-delimiter"
                                            type="text"
                                            value={outputConfig.delimiter || ''}
                                            onChange={handleDelimiterChange}
                                            placeholder="Ex: |"
                                            className="w-full"
                                            disabled={isProcessing}
                                            maxLength={5}
                                        />
                                    </div>
                                )}
                            </div>

                             <div>
                                 <h3 className="text-lg font-medium mb-2">Campos de Saída</h3>
                                  <p className="text-xs text-muted-foreground mb-2">Defina a ordem, conteúdo e formatação dos campos no arquivo final. Use os botões <ArrowUp className='inline h-3 w-3'/> / <ArrowDown className='inline h-3 w-3'/> para reordenar.</p>
                                 <div className="max-h-[45vh] overflow-auto border rounded-md">
                                     <Table>
                                         <TableHeader>
                                             <TableRow>
                                                  <TableHead className="w-[70px]">Ordem</TableHead>
                                                  <TableHead className="w-3/12">Campo</TableHead>
                                                   <TableHead className="w-2/12">Formato Data</TableHead>
                                                  {outputConfig.format === 'txt' && (<>
                                                          <TableHead className="w-[80px]">
                                                              Tam.
                                                              <TooltipProvider>
                                                                  <Tooltip>
                                                                      <TooltipTrigger asChild><Button variant="ghost" size="icon" className="ml-1 h-6 w-6 text-muted-foreground cursor-help align-middle"><HelpCircle className="h-4 w-4" /></Button></TooltipTrigger>
                                                                      <TooltipContent><p>Tamanho fixo (obrigatório).</p></TooltipContent>
                                                                  </Tooltip>
                                                              </TooltipProvider>
                                                          </TableHead>
                                                           <TableHead className="w-[80px]">
                                                              Preench.
                                                               <TooltipProvider>
                                                                  <Tooltip>
                                                                      <TooltipTrigger asChild><Button variant="ghost" size="icon" className="ml-1 h-6 w-6 text-muted-foreground cursor-help align-middle"><HelpCircle className="h-4 w-4" /></Button></TooltipTrigger>
                                                                      <TooltipContent><p>Caractere (1) p/ preencher.</p></TooltipContent>
                                                                  </Tooltip>
                                                              </TooltipProvider>
                                                           </TableHead>
                                                           <TableHead className="w-2/12">
                                                              Direção Preench.
                                                               <TooltipProvider>
                                                                  <Tooltip>
                                                                      <TooltipTrigger asChild><Button variant="ghost" size="icon" className="ml-1 h-6 w-6 text-muted-foreground cursor-help align-middle"><HelpCircle className="h-4 w-4" /></Button></TooltipTrigger>
                                                                      <TooltipContent>
                                                                            <p>Esquerda (p/ números) ou Direita (p/ texto).</p>
                                                                       </TooltipContent>
                                                                  </Tooltip>
                                                              </TooltipProvider>
                                                           </TableHead>
                                                  </>)}
                                                  <TableHead className={`w-[100px] text-right ${outputConfig.format === 'csv' ? 'pl-20' : ''}`}>Ações</TableHead>
                                             </TableRow>
                                         </TableHeader>
                                         <TableBody>
                                             {outputConfig.fields.map((field, index) => {
                                                 const dataType = getOutputFieldDataType(field);
                                                 const isDateField = dataType === 'Data';
                                                 let fieldNameDisplayElement: React.ReactNode;

                                                 if (field.isStatic && field.isStaticFromPreset) {
                                                      fieldNameDisplayElement = (
                                                          <TooltipProvider>
                                                              <Tooltip>
                                                                  <TooltipTrigger asChild>
                                                                      <span className="font-medium text-blue-600 dark:text-blue-400 cursor-help underline-dashed">
                                                                          {field.fieldName}
                                                                      </span>
                                                                  </TooltipTrigger>
                                                                  <TooltipContent>
                                                                      <p>Valor Padrão (do Layout): {field.staticValue}</p>
                                                                  </TooltipContent>
                                                              </Tooltip>
                                                          </TooltipProvider>
                                                      );
                                                  } else if (field.isStatic) {
                                                     fieldNameDisplayElement = <span className="font-medium text-blue-600 dark:text-blue-400" title={`Valor: ${field.staticValue}`}>{field.fieldName}</span>;
                                                 } else if (field.isCalculated) {
                                                      fieldNameDisplayElement = <span className="font-medium text-purple-600 dark:text-purple-400" title={`Tipo: ${field.type}`}>{field.fieldName}</span>;
                                                 } else {
                                                      fieldNameDisplayElement = renderMappedOutputFieldSelect(field);
                                                 }


                                                 return (
                                                 <TableRow key={field.id}>
                                                      <TableCell className="flex items-center gap-1">
                                                         <span className="text-xs w-6 text-center">{index + 1}</span>
                                                         <div className='flex flex-col'>
                                                            <Button variant="ghost" size="icon" className="h-5 w-5" onClick={() => moveField(field.id, 'up')} disabled={isProcessing || index === 0} aria-label="Mover para cima">
                                                                 <ArrowUp className="h-3 w-3" />
                                                             </Button>
                                                             <Button variant="ghost" size="icon" className="h-5 w-5" onClick={() => moveField(field.id, 'down')} disabled={isProcessing || index === outputConfig.fields.length - 1} aria-label="Mover para baixo">
                                                                 <ArrowDown className="h-3 w-3" />
                                                             </Button>
                                                         </div>
                                                      </TableCell>
                                                     <TableCell className="text-xs">
                                                         {field.isStatic || field.isCalculated ? (
                                                             <div className="flex items-center gap-1">
                                                                {fieldNameDisplayElement}
                                                                <Button variant="ghost" size="icon" className="h-6 w-6 text-muted-foreground hover:text-accent" onClick={() => field.isStatic ? openEditStaticFieldDialog(field) : openEditCalculatedFieldDialog(field)} aria-label={`Editar campo ${field.fieldName}`}>
                                                                     <Edit className="h-3 w-3" />
                                                                 </Button>
                                                             </div>
                                                         ) : (
                                                            fieldNameDisplayElement
                                                         )}
                                                     </TableCell>
                                                     <TableCell>
                                                          <Select
                                                               value={field.dateFormat ?? ''}
                                                               onValueChange={(value) => handleOutputFieldChange(field.id, 'dateFormat', value)}
                                                               disabled={isProcessing || !isDateField}
                                                            >
                                                                <SelectTrigger className={`w-full h-8 text-xs ${!isDateField ? 'invisible' : ''}`}>
                                                                    <SelectValue placeholder="Formato Data" />
                                                                </SelectTrigger>
                                                                <SelectContent>
                                                                     <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>-- Selecione --</SelectItem>
                                                                    {DATE_FORMATS.map(df => (
                                                                         <SelectItem key={df} value={df}>{df === 'YYYYMMDD' ? 'AAAAMMDD' : 'DDMMAAAA'}</SelectItem>
                                                                    ))}
                                                                </SelectContent>
                                                          </Select>
                                                     </TableCell>
                                                     {outputConfig.format === 'txt' && (<>
                                                          <TableCell>
                                                             <Input
                                                                 type="number"
                                                                 min="1"
                                                                 value={field.length ?? ''}
                                                                 onChange={(e) => handleOutputFieldChange(field.id, 'length', e.target.value)}
                                                                 placeholder="Obrig."
                                                                 className="w-full h-8 text-xs"
                                                                 required
                                                                 disabled={isProcessing}
                                                                 aria-label={`Tamanho do campo ${field.isStatic || field.isCalculated ? field.fieldName : field.mappedField}`}
                                                             />
                                                          </TableCell>
                                                          <TableCell>
                                                             <Input
                                                                type="text"
                                                                maxLength={1}
                                                                value={field.paddingChar ?? ''}
                                                                onChange={(e) => handleOutputFieldChange(field.id, 'paddingChar', e.target.value)}
                                                                 placeholder={getDefaultPaddingChar(field, columnMappings)}
                                                                className="w-10 text-center h-8 text-xs"
                                                                 required
                                                                disabled={isProcessing}
                                                                aria-label={`Caractere de preenchimento do campo ${field.isStatic || field.isCalculated ? field.fieldName : field.mappedField}`}
                                                             />
                                                         </TableCell>
                                                         <TableCell>
                                                              <Select
                                                                  value={field.paddingDirection ?? getDefaultPaddingDirection(field, columnMappings)}
                                                                 onValueChange={(value) => handleOutputFieldChange(field.id, 'paddingDirection', value)}
                                                                 disabled={isProcessing}
                                                               >
                                                                  <SelectTrigger className="w-full h-8 text-xs">
                                                                       <SelectValue />
                                                                   </SelectTrigger>
                                                                   <SelectContent>
                                                                        <SelectItem value="left">Esquerda</SelectItem>
                                                                        <SelectItem value="right">Direita</SelectItem>
                                                                    </SelectContent>
                                                              </Select>
                                                          </TableCell>
                                                        </>)}
                                                      <TableCell className={`text-right ${outputConfig.format === 'csv' ? 'pl-20' : ''}`}>
                                                          <TooltipProvider>
                                                              <Tooltip>
                                                                  <TooltipTrigger asChild>
                                                                         <Button
                                                                             variant="ghost"
                                                                             size="icon"
                                                                             onClick={() => removeOutputField(field.id)}
                                                                             disabled={isProcessing}
                                                                             className="h-7 w-7 text-muted-foreground hover:text-destructive"
                                                                             aria-label={`Remover campo ${field.isStatic || field.isCalculated ? field.fieldName : field.mappedField} da saída`}
                                                                         >
                                                                             <Trash2 className="h-4 w-4" />
                                                                         </Button>
                                                                   </TooltipTrigger>
                                                                   <TooltipContent>
                                                                       <p>Remover campo da saída</p>
                                                                   </TooltipContent>
                                                               </Tooltip>
                                                           </TooltipProvider>
                                                     </TableCell>
                                                 </TableRow>
                                                 );
                                            })}
                                             {outputConfig.fields.length === 0 && (
                                                 <TableRow>
                                                      <TableCell colSpan={outputConfig.format === 'txt' ? 7 : 4} className="text-center text-muted-foreground py-4">
                                                         Nenhum campo adicionado à saída. Use os botões abaixo.
                                                     </TableCell>
                                                 </TableRow>
                                              )}
                                         </TableBody>
                                     </Table>
                                  </div>
                                   <div className="flex gap-2 mt-2">
                                      <Button onClick={addOutputField} variant="outline" size="sm" disabled={isProcessing || columnMappings.filter(m => m.mappedField !== null && !outputConfig.fields.some(of => !of.isStatic && !of.isCalculated && of.mappedField === m.mappedField)).length === 0}>
                                          <Plus className="mr-2 h-4 w-4" /> Adicionar Campo Mapeado
                                      </Button>
                                      <Button onClick={openAddStaticFieldDialog} variant="outline" size="sm" disabled={isProcessing}>
                                          <Plus className="mr-2 h-4 w-4" /> Adicionar Campo Estático
                                      </Button>
                                       <Button onClick={openAddCalculatedFieldDialog} variant="outline" size="sm" disabled={isProcessing}>
                                            <Calculator className="mr-2 h-4 w-4" /> Adicionar Campo Calculado
                                       </Button>
                                   </div>
                             </div>
                         </CardContent>
                         <CardFooter className="flex justify-between">
                             <Button variant="outline" onClick={() => setActiveTab("mapping")} disabled={isProcessing}>Voltar</Button>
                             <Button onClick={convertFile} disabled={isProcessing || outputConfig.fields.length === 0} variant="default">
                                 {isProcessing ? (
                                    <>
                                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                                        Convertendo...
                                    </>
                                    ) : (
                                    <>
                                        Iniciar Conversão <ArrowRight className="ml-2 h-4 w-4" />
                                    </>
                                    )}

                             </Button>
                         </CardFooter>
                    </Card>
                 </div>
                )}
                 {!isProcessing && !file && (
                    <p className="text-center text-muted-foreground p-4">Complete as etapas de Upload e Mapeamento primeiro.</p>
                )}
            </TabsContent>

             {/* 4. Result Tab */}
            <TabsContent value="result">
               {isProcessing && activeTab === "result" && (
                 <div className="flex items-center justify-center text-accent animate-pulse p-4">
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      {processingMessage}
                  </div>
               )}
                {!isProcessing && convertedData && (
                    <Card>
                         <CardHeader>
                             <CardTitle className="text-xl">Resultado da Conversão</CardTitle>
                             <CardDescription>
                                Pré-visualização do arquivo convertido ({outputConfig.format.toUpperCase()}, {outputConfig.encoding}). Verifique antes de baixar.
                             </CardDescription>
                         </CardHeader>
                         <CardContent>
                             <Textarea
                                 readOnly
                                 value={convertedData instanceof Buffer
                                          ? iconv.decode(convertedData, outputConfig.encoding)
                                          : String(convertedData) }
                                 className="w-full h-64 font-mono text-xs bg-secondary/30 border rounded-md"
                                 placeholder="Resultado da conversão aparecerá aqui..."
                                 aria-label="Pré-visualização do arquivo convertido"
                             />
                         </CardContent>
                         <CardFooter className="flex flex-col sm:flex-row justify-between gap-2">
                             <Button variant="outline" onClick={() => setActiveTab("config")} disabled={isProcessing}>Voltar à Configuração</Button>
                            <div className="flex gap-2">
                                 <Button onClick={resetState} variant="outline" className="mr-2" disabled={isProcessing}>
                                     <RotateCcw className="mr-2 h-4 w-4" /> Nova Conversão
                                 </Button>
                                 <Button onClick={openDownloadDialog} disabled={isProcessing || !convertedData} variant="default">
                                     <Download className="mr-2 h-4 w-4" /> Baixar Arquivo Convertido
                                 </Button>
                            </div>
                         </CardFooter>
                    </Card>
                )}
                 {!isProcessing && !convertedData && (
                    <p className="text-center text-muted-foreground p-4">Execute a conversão na aba "Configurar Saída" para ver o resultado.</p>
                )}
            </TabsContent>
          </Tabs>
        </CardContent>

        <CardFooter className="text-xs text-muted-foreground pt-4 border-t flex flex-col sm:flex-row justify-between items-center gap-2">
             <span className="text-left">
                 © {new Date().getFullYear()} SCA. Ferramenta de conversão de dados. - Desenvolvido por <a href="mailto:faraujo@gmail.com" className="text-accent hover:underline">Fábio Araújo</a>
             </span>
             <span className="font-mono text-accent text-right">v{appVersion}</span>
        </CardFooter>
      </Card>

        {/* Add/Edit Static Field Dialog */}
        <Dialog open={staticFieldDialogState.isOpen} onOpenChange={(isOpen) => setStaticFieldDialogState(prev => ({ ...prev, isOpen }))}>
            <DialogContent className="sm:max-w-[425px]">
                <DialogHeader>
                    <DialogTitle>{staticFieldDialogState.isEditing ? 'Editar' : 'Adicionar'} Campo Estático</DialogTitle>
                    <DialogDescription>
                       Defina um campo com valor fixo para incluir no arquivo de saída.
                    </DialogDescription>
                </DialogHeader>
                 <div className="grid gap-4 py-4">
                    <div className="space-y-2">
                        <Label htmlFor="static-field-name">Nome*</Label>
                        <Input
                            id="static-field-name"
                            value={staticFieldDialogState.fieldName}
                            onChange={(e) => handleStaticFieldDialogChange('fieldName', e.target.value)}
                            placeholder="Ex: FlagAtivo"
                            required
                        />
                    </div>
                    <div className="space-y-2">
                        <Label htmlFor="static-field-value">Valor</Label>
                        <Input
                            id="static-field-value"
                            value={staticFieldDialogState.staticValue}
                            onChange={(e) => handleStaticFieldDialogChange('staticValue', e.target.value)}
                            placeholder="Ex: S ou 001001"
                        />
                    </div>
                     {outputConfig.format === 'txt' && (
                         <>
                            <div className="space-y-2">
                                <Label htmlFor="static-field-length">Tamanho* (TXT)</Label>
                                <Input
                                    id="static-field-length"
                                    type="number"
                                    min="1"
                                    value={staticFieldDialogState.length}
                                    onChange={(e) => handleStaticFieldDialogChange('length', e.target.value)}
                                    required
                                    placeholder="Obrigatório para TXT"
                                />
                             </div>
                            <div className="grid grid-cols-2 gap-4">
                                <div className="space-y-2">
                                    <Label htmlFor="static-field-padding-char">Preencher* (TXT)</Label>
                                    <Input
                                        id="static-field-padding-char"
                                        type="text"
                                        maxLength={1}
                                        value={staticFieldDialogState.paddingChar}
                                        onChange={(e) => handleStaticFieldDialogChange('paddingChar', e.target.value)}
                                        className="text-center"
                                        required
                                        placeholder={/^-?\d+$/.test(staticFieldDialogState.staticValue) ? '0' : ' '}
                                    />
                                </div>
                                <div className="space-y-2">
                                    <Label htmlFor="static-field-padding-direction">Direção* (TXT)</Label>
                                     <Select
                                           value={staticFieldDialogState.paddingDirection}
                                           onValueChange={(value) => handleStaticFieldDialogChange('paddingDirection', value)}
                                           disabled={isProcessing}
                                        >
                                           <SelectTrigger id="static-field-padding-direction">
                                                <SelectValue />
                                            </SelectTrigger>
                                            <SelectContent>
                                                 <SelectItem value="left">Esquerda</SelectItem>
                                                 <SelectItem value="right">Direita</SelectItem>
                                             </SelectContent>
                                      </Select>
                                  </div>
                             </div>
                         </>
                     )}
                      <p className="text-xs text-muted-foreground">* Campos obrigatórios.</p>
                </div>
                <DialogFooter>
                    <DialogClose asChild>
                        <Button type="button" variant="outline">Cancelar</Button>
                    </DialogClose>
                    <Button type="button" onClick={saveStaticField}>Salvar Campo</Button>
                </DialogFooter>
            </DialogContent>
        </Dialog>

        {/* Add/Edit Predefined Field Dialog */}
        <Dialog open={predefinedFieldDialogState.isOpen} onOpenChange={(isOpen) => setPredefinedFieldDialogState(prev => ({ ...prev, isOpen }))}>
             <DialogContent className="sm:max-w-[425px]">
                 <DialogHeader>
                     <DialogTitle>{predefinedFieldDialogState.isEditing ? 'Editar' : 'Adicionar'} Campo Pré-definido</DialogTitle>
                     <DialogDescription>
                         {predefinedFieldDialogState.isEditing
                             ? "Edite as propriedades do campo pré-definido."
                             : "Adicione um novo campo para usar no mapeamento."}
                          Defina se o campo será Principal (mantido para futuras conversões) ou Opcional.
                     </DialogDescription>
                 </DialogHeader>
                 <div className="grid gap-4 py-4">
                     <div className="space-y-2">
                         <Label htmlFor="predefined-field-name">Nome*</Label>
                         <Input
                             id="predefined-field-name"
                             value={predefinedFieldDialogState.fieldName}
                             onChange={(e) => handlePredefinedFieldDialogChange('fieldName', e.target.value)}
                             placeholder="Ex: Código do Cliente"
                             required
                         />
                     </div>
                     <div className="space-y-2">
                         <Label htmlFor="predefined-field-comment">Comentário</Label>
                         <Textarea
                             id="predefined-field-comment"
                             value={predefinedFieldDialogState.comment}
                             onChange={(e) => handlePredefinedFieldDialogChange('comment', e.target.value)}
                             className="min-h-[60px]"
                             placeholder="Opcional: Descrição curta ou instrução de uso (ex: Usar apenas números)"
                         />
                     </div>
                     <div className="flex items-center space-x-2 pt-2">
                         <Checkbox
                             id="predefined-persist"
                             checked={predefinedFieldDialogState.isPersistent}
                             onCheckedChange={(checked) => handlePredefinedFieldDialogChange('isPersistent', Boolean(checked))}
                             aria-label="Marcar como Campo Principal"
                             disabled={predefinedFields.find(f => f.id === predefinedFieldDialogState.fieldId)?.isCore}
                         />
                         <Label htmlFor="predefined-persist" className="cursor-pointer">
                             Campo Principal (Manter para futuras conversões)
                         </Label>
                          <TooltipProvider>
                               <Tooltip>
                                   <TooltipTrigger asChild>
                                        <HelpCircle className="h-4 w-4 text-muted-foreground cursor-help" />
                                   </TooltipTrigger>
                                   <TooltipContent>
                                       <p>Campos Principais são salvos no seu navegador e ficam disponíveis para todas as conversões.</p>
                                       <p>Campos Opcionais são usados apenas nesta conversão.</p>
                                       <p>(Campos originais são sempre Principais).</p>
                                   </TooltipContent>
                               </Tooltip>
                           </TooltipProvider>
                     </div>
                     <p className="text-xs text-muted-foreground">* Nome é obrigatório.</p>
                 </div>
                 <DialogFooter>
                     <DialogClose asChild>
                         <Button type="button" variant="outline">Cancelar</Button>
                     </DialogClose>
                     <Button type="button" onClick={savePredefinedField}>
                         {predefinedFieldDialogState.isEditing ? 'Salvar Alterações' : 'Adicionar Campo'}
                     </Button>
                 </DialogFooter>
             </DialogContent>
         </Dialog>

        {/* Add/Edit Calculated Field Dialog */}
        <Dialog open={calculatedFieldDialogState.isOpen} onOpenChange={(isOpen) => setCalculatedFieldDialogState(prev => ({ ...prev, isOpen }))}>
            <DialogContent className="sm:max-w-[500px]">
                <DialogHeader>
                    <DialogTitle>{calculatedFieldDialogState.isEditing ? 'Editar' : 'Adicionar'} Campo Calculado</DialogTitle>
                    <DialogDescription>
                        Defina um campo cujo valor é calculado com base em outros campos mapeados ou parâmetros.
                    </DialogDescription>
                </DialogHeader>
                <div className="grid gap-4 py-4 max-h-[60vh] overflow-y-auto pr-2">
                    <div className="space-y-2">
                        <Label htmlFor="calc-field-name">Nome do Campo*</Label>
                        <Input
                            id="calc-field-name"
                            value={calculatedFieldDialogState.fieldName}
                            onChange={(e) => handleCalculatedFieldDialogChange('fieldName', e.target.value)}
                            placeholder="Ex: Data Início Contrato"
                            required
                        />
                    </div>

                    <div className="space-y-2">
                        <Label htmlFor="calc-field-type">Tipo de Cálculo*</Label>
                        <Select
                            value={calculatedFieldDialogState.type || NONE_VALUE_PLACEHOLDER}
                            onValueChange={(value) => handleCalculatedFieldDialogChange('type', value)}
                            required
                        >
                            <SelectTrigger id="calc-field-type">
                                <SelectValue placeholder="Selecione o tipo..." />
                            </SelectTrigger>
                            <SelectContent>
                                <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>-- Selecione --</SelectItem>
                                <SelectItem value="CalculateStartDate">Calcular Data Inicial (Periodo - Parc. Pagas)</SelectItem>
                                <SelectItem value="CalculateSituacaoRetorno">Calcular Situação Retorno (Valor Realizado)</SelectItem>
                                <SelectItem value="FormatPeriodMMAAAA">Formatar Período para MMAAAA (do campo 'Período' ou Data Atual)</SelectItem>
                            </SelectContent>
                        </Select>
                    </div>

                    {calculatedFieldDialogState.type === 'CalculateStartDate' && (
                        <>
                            <div className="space-y-2 p-3 border rounded-md bg-muted/50">
                                <h4 className="text-sm font-medium mb-2">Parâmetros para "Calcular Data Inicial"</h4>
                                <div className="space-y-2">
                                    <Label htmlFor="calc-param-period">Período Atual* (DD/MM/AAAA)</Label>
                                    <Input
                                        id="calc-param-period"
                                        value={calculatedFieldDialogState.parameters.period || ''}
                                        onChange={(e) => handleCalculatedFieldDialogChange('parameters.period', e.target.value)}
                                        placeholder="Ex: 31/12/2024"
                                        required
                                    />
                                </div>
                                <div className="space-y-2">
                                    <Label htmlFor="calc-req-parcelas">Campo Mapeado "Parcelas Pagas"*</Label>
                                    <Select
                                        value={calculatedFieldDialogState.requiredInputFields.parcelasPagas || NONE_VALUE_PLACEHOLDER}
                                        onValueChange={(value) => handleCalculatedFieldDialogChange('requiredInputFields.parcelasPagas', value)}
                                        required
                                    >
                                        <SelectTrigger id="calc-req-parcelas">
                                            <SelectValue placeholder="Selecione o campo..." />
                                        </SelectTrigger>
                                        <SelectContent>
                                            <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>-- Selecione --</SelectItem>
                                            {columnMappings
                                                .filter(m => m.mappedField && (m.dataType === 'Inteiro' || m.dataType === 'Numérico'))
                                                .map(m => {
                                                    const predefined = predefinedFields.find(pf => pf.id === m.mappedField);
                                                    return (
                                                        <SelectItem key={m.mappedField!} value={m.mappedField!}>
                                                            {predefined?.name ?? m.mappedField} (Coluna: {m.originalHeader})
                                                        </SelectItem>
                                                    );
                                                })}
                                             {columnMappings.filter(m => m.mappedField && !(m.dataType === 'Inteiro' || m.dataType === 'Numérico')).length > 0 && (
                                                 <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>--- Outros Campos Mapeados ---</SelectItem>
                                             )}
                                              {columnMappings
                                                .filter(m => m.mappedField && !(m.dataType === 'Inteiro' || m.dataType === 'Numérico'))
                                                .map(m => {
                                                    const predefined = predefinedFields.find(pf => pf.id === m.mappedField);
                                                    return (
                                                        <SelectItem key={m.mappedField!} value={m.mappedField!}>
                                                            {predefined?.name ?? m.mappedField} (Coluna: {m.originalHeader})
                                                        </SelectItem>
                                                    );
                                                })}
                                        </SelectContent>
                                    </Select>
                                    <p className='text-xs text-muted-foreground'>Selecione o campo que contém o número de parcelas pagas.</p>
                                </div>
                                <div className="space-y-2">
                                     <Label htmlFor="calc-date-format">Formato da Data de Saída*</Label>
                                     <Select
                                         value={calculatedFieldDialogState.dateFormat || NONE_VALUE_PLACEHOLDER}
                                         onValueChange={(value) => handleCalculatedFieldDialogChange('dateFormat', value)}
                                         required
                                     >
                                         <SelectTrigger id="calc-date-format">
                                             <SelectValue placeholder="Selecione..." />
                                         </SelectTrigger>
                                         <SelectContent>
                                             <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>-- Selecione --</SelectItem>
                                              {DATE_FORMATS.map(df => (
                                                   <SelectItem key={df} value={df}>{df === 'YYYYMMDD' ? 'AAAAMMDD' : 'DDMMAAAA'}</SelectItem>
                                              ))}
                                         </SelectContent>
                                     </Select>
                                 </div>
                            </div>
                        </>
                    )}
                     {calculatedFieldDialogState.type === 'CalculateSituacaoRetorno' && (
                        <div className="space-y-2 p-3 border rounded-md bg-muted/50">
                            <h4 className="text-sm font-medium mb-2">Parâmetros para "Calcular Situação Retorno"</h4>
                            <div className="space-y-2">
                                <Label htmlFor="calc-req-valor-realizado">Campo Mapeado "Valor Realizado"*</Label>
                                <Select
                                    value={calculatedFieldDialogState.requiredInputFields.valorRealizado || NONE_VALUE_PLACEHOLDER}
                                    onValueChange={(value) => handleCalculatedFieldDialogChange('requiredInputFields.valorRealizado', value)}
                                    required
                                >
                                    <SelectTrigger id="calc-req-valor-realizado">
                                        <SelectValue placeholder="Selecione o campo..." />
                                    </SelectTrigger>
                                    <SelectContent>
                                        <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>-- Selecione --</SelectItem>
                                         {columnMappings
                                            .filter(m => m.mappedField && m.dataType === 'Numérico')
                                            .map(m => {
                                                const predefined = predefinedFields.find(pf => pf.id === m.mappedField);
                                                return (
                                                    <SelectItem key={m.mappedField!} value={m.mappedField!}>
                                                        {predefined?.name ?? m.mappedField} (Coluna: {m.originalHeader})
                                                    </SelectItem>
                                                );
                                         })}
                                        {columnMappings.filter(m => m.mappedField && m.dataType !== 'Numérico').length > 0 && (
                                              <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>--- Outros Campos Mapeados ---</SelectItem>
                                          )}
                                        {columnMappings
                                            .filter(m => m.mappedField && m.dataType !== 'Numérico')
                                            .map(m => {
                                                const predefined = predefinedFields.find(pf => pf.id === m.mappedField);
                                                return (
                                                    <SelectItem key={m.mappedField!} value={m.mappedField!}>
                                                        {predefined?.name ?? m.mappedField} (Coluna: {m.originalHeader})
                                                    </SelectItem>
                                                );
                                         })}
                                    </SelectContent>
                                </Select>
                                <p className='text-xs text-muted-foreground'>Selecione o campo que contém o valor realizado.</p>
                            </div>
                        </div>
                    )}
                     {calculatedFieldDialogState.type === 'FormatPeriodMMAAAA' && (
                         <div className="space-y-2 p-3 border rounded-md bg-muted/50">
                            <h4 className="text-sm font-medium mb-2">Parâmetros para "Formatar Período para MMAAAA"</h4>
                             <p className="text-xs text-muted-foreground">
                                 Este cálculo usará o campo "Período" (se mapeado na Aba 2 e com tipo Data) ou a data atual para gerar o formato MMAAAA.
                                 O formato de saída da data será MMAAAA (ex: 052024).
                             </p>
                             {/* No specific parameters needed here beyond the optional mapped "Período" field */}
                         </div>
                     )}


                     {outputConfig.format === 'txt' && (
                         <div className="p-3 border rounded-md bg-muted/50 space-y-4">
                              <h4 className="text-sm font-medium mb-2">Formatação de Saída (TXT)</h4>
                              <div className="space-y-2">
                                 <Label htmlFor="calc-field-length">Tamanho* (TXT)</Label>
                                 <Input
                                     id="calc-field-length"
                                     type="number"
                                     min="1"
                                     value={calculatedFieldDialogState.length}
                                     onChange={(e) => handleCalculatedFieldDialogChange('length', e.target.value)}
                                     required
                                     placeholder="Obrigatório para TXT"
                                 />
                              </div>
                             <div className="grid grid-cols-2 gap-4">
                                 <div className="space-y-2">
                                     <Label htmlFor="calc-field-padding-char">Preencher* (TXT)</Label>
                                     <Input
                                         id="calc-field-padding-char"
                                         type="text"
                                         maxLength={1}
                                         value={calculatedFieldDialogState.paddingChar}
                                         onChange={(e) => handleCalculatedFieldDialogChange('paddingChar', e.target.value)}
                                         className="text-center"
                                         required
                                          placeholder={getDefaultPaddingChar({ isCalculated: true, id: '', order: 0, type: calculatedFieldDialogState.type || 'CalculateStartDate', fieldName: '', requiredInputFields: [] }, columnMappings)}
                                     />
                                 </div>
                                 <div className="space-y-2">
                                     <Label htmlFor="calc-field-padding-direction">Direção* (TXT)</Label>
                                      <Select
                                            value={calculatedFieldDialogState.paddingDirection}
                                            onValueChange={(value) => handleCalculatedFieldDialogChange('paddingDirection', value)}
                                         >
                                            <SelectTrigger id="calc-field-padding-direction">
                                                 <SelectValue />
                                             </SelectTrigger>
                                             <SelectContent>
                                                  <SelectItem value="left">Esquerda</SelectItem>
                                                  <SelectItem value="right">Direita</SelectItem>
                                              </SelectContent>
                                       </Select>
                                   </div>
                              </div>
                         </div>
                     )}
                      <p className="text-xs text-muted-foreground">* Campos obrigatórios.</p>
                </div>
                <DialogFooter>
                    <DialogClose asChild>
                        <Button type="button" variant="outline">Cancelar</Button>
                    </DialogClose>
                    <Button type="button" onClick={saveCalculatedField}>Salvar Campo Calculado</Button>
                </DialogFooter>
            </DialogContent>
        </Dialog>


       {/* Download File Dialog */}
        <Dialog open={downloadDialogState.isOpen} onOpenChange={(isOpen) => setDownloadDialogState(prev => ({ ...prev, isOpen }))}>
            <DialogContent className="sm:max-w-[425px]">
                <DialogHeader>
                    <DialogTitle>Renomear e Baixar Arquivo</DialogTitle>
                    <DialogDescription>
                       Confirme ou altere o nome do arquivo antes de baixar.
                    </DialogDescription>
                </DialogHeader>
                <div className="space-y-4 py-4">
                   <div className="space-y-2">
                        <Label htmlFor="download-filename">Nome do Arquivo*</Label>
                        <Input
                            id="download-filename"
                            value={downloadDialogState.finalFilename}
                            onChange={handleDownloadFilenameChange}
                            placeholder="Nome do arquivo de saída"
                            required
                        />
                   </div>
                    <p className="text-xs text-muted-foreground">A extensão (.txt ou .csv) será adicionada automaticamente se ausente.</p>
                    <p className="text-xs text-muted-foreground">* Nome do arquivo é obrigatório.</p>
                </div>
                <DialogFooter>
                    <DialogClose asChild>
                        <Button type="button" variant="outline">Cancelar</Button>
                    </DialogClose>
                    <Button type="button" onClick={confirmDownload} disabled={!downloadDialogState.finalFilename.trim()}>Confirmar Download</Button>
                </DialogFooter>
            </DialogContent>
        </Dialog>

        <Dialog open={configManagementDialogState.isOpen} onOpenChange={(isOpen) => setConfigManagementDialogState(prev => ({ ...prev, isOpen }))}>
             <DialogContent className="sm:max-w-[425px]">
                 <DialogHeader>
                     <DialogTitle>{configManagementDialogState.action === 'save' ? 'Salvar' : 'Carregar'} Modelo de Configuração</DialogTitle>
                     <DialogDescription>
                         {configManagementDialogState.action === 'save'
                             ? 'Digite um nome para salvar a configuração atual (formato, campos, etc.). Se o nome já existir, ele será sobrescrito.'
                             : 'Selecione um modelo salvo para carregar suas configurações.'}
                     </DialogDescription>
                 </DialogHeader>
                 <div className="grid gap-4 py-4">
                     {configManagementDialogState.action === 'save' && (
                         <div className="space-y-2">
                             <Label htmlFor="config-name">Nome do Modelo*</Label>
                             <Input
                                 id="config-name"
                                 value={configManagementDialogState.configName}
                                 onChange={(e) => handleConfigDialogChange('configName', e.target.value)}
                                 placeholder="Ex: Layout Banco X v1"
                                 required
                             />
                         </div>
                     )}
                     {configManagementDialogState.action === 'load' && (
                         <div className="space-y-2">
                             <Label htmlFor="config-load-select">Selecionar Modelo*</Label>
                             <Select
                                 value={configManagementDialogState.selectedConfigToLoad || NONE_VALUE_PLACEHOLDER}
                                 onValueChange={(value) => handleConfigDialogChange('selectedConfigToLoad', value)}
                                 required
                             >
                                 <SelectTrigger id="config-load-select">
                                     <SelectValue placeholder="Selecione um modelo..." />
                                 </SelectTrigger>
                                 <SelectContent>
                                     <SelectItem value={NONE_VALUE_PLACEHOLDER} disabled>-- Selecione --</SelectItem>
                                     {savedConfigs.map(config => (
                                          <SelectItem key={config.name} value={config.name!}>
                                             <div className="flex justify-between items-center w-full">
                                                  <span>{config.name}</span>
                                                  <Popover>
                                                      <PopoverTrigger asChild onClick={(e) => e.stopPropagation()}>
                                                            <Button
                                                                variant="ghost"
                                                                size="icon"
                                                                className="h-5 w-5 ml-2 text-muted-foreground hover:text-accent shrink-0"
                                                                aria-label={`Opções para ${config.name}`}
                                                                onPointerDown={(e) => e.stopPropagation()}
                                                            >
                                                                <Settings className="h-3 w-3"/>
                                                            </Button>
                                                      </PopoverTrigger>
                                                      <PopoverContent className="w-auto p-1" onClick={(e) => e.stopPropagation()}>
                                                          <Button
                                                              variant="ghost"
                                                              size="sm"
                                                              className="w-full justify-start text-xs h-7"
                                                              onClick={(e) => {
                                                                  e.stopPropagation();
                                                                  const jsonString = JSON.stringify(config, null, 2);
                                                                  const blob = new Blob([jsonString], { type: "application/json" });
                                                                  const url = URL.createObjectURL(blob);
                                                                  const a = document.createElement("a");
                                                                  a.href = url;
                                                                  a.download = `${config.name?.replace(/\s+/g, '_') || 'modelo'}_sca.json`;
                                                                  document.body.appendChild(a);
                                                                  a.click();
                                                                  document.body.removeChild(a);
                                                                  URL.revokeObjectURL(url);
                                                                  toast({ title: "Modelo Baixado", description: `O modelo "${config.name}" foi baixado.`});
                                                              }}
                                                          >
                                                              <Download className="mr-2 h-3 w-3" /> Baixar JSON
                                                          </Button>
                                                           <Button
                                                              variant="ghost"
                                                              size="sm"
                                                              className="w-full justify-start text-xs text-destructive hover:text-destructive h-7"
                                                              onClick={(e) => { e.stopPropagation(); deleteConfig(config.name!); }}
                                                          >
                                                              <Trash2 className="mr-2 h-3 w-3" /> Excluir
                                                          </Button>
                                                      </PopoverContent>
                                                  </Popover>
                                              </div>
                                         </SelectItem>
                                     ))}
                                     {savedConfigs.length === 0 && <SelectItem value="no-configs" disabled>Nenhum modelo salvo.</SelectItem>}
                                 </SelectContent>
                             </Select>
                              <div className="mt-2">
                                <Label htmlFor="config-upload-json" className="text-xs text-muted-foreground">Ou carregue um modelo .json:</Label>
                                <Input
                                  id="config-upload-json"
                                  type="file"
                                  accept=".json"
                                  className="mt-1 text-xs h-9"
                                  onChange={(e) => {
                                    const file = e.target.files?.[0];
                                    if (file) {
                                      const reader = new FileReader();
                                      reader.onload = (event) => {
                                        try {
                                          const loadedConfig = JSON.parse(event.target?.result as string) as OutputConfig;
                                          if (loadedConfig.name && Array.isArray(loadedConfig.fields)) {
                                            setOutputConfig(loadedConfig);
                                            setSelectedLayoutPresetId("custom");
                                            setSavedConfigs(prev => {
                                                const existing = prev.find(c => c.name === loadedConfig.name);
                                                let updated;
                                                if (existing) {
                                                    updated = prev.map(c => c.name === loadedConfig.name ? loadedConfig : c);
                                                } else {
                                                    updated = [...prev, loadedConfig];
                                                }
                                                saveAllConfigs(updated);
                                                return updated;
                                            });
                                            setConfigManagementDialogState({ isOpen: false, action: null, configName: '', selectedConfigToLoad: null });
                                            toast({ title: "Sucesso", description: `Modelo "${loadedConfig.name}" carregado do arquivo.` });
                                          } else {
                                            toast({ title: "Erro", description: "Arquivo JSON inválido ou não corresponde ao formato esperado.", variant: "destructive" });
                                          }
                                        } catch (error) {
                                          toast({ title: "Erro ao ler JSON", description: "Não foi possível processar o arquivo JSON.", variant: "destructive" });
                                          console.error("Error parsing JSON config:", error);
                                        }
                                      };
                                      reader.readAsText(file);
                                      if (e.target) e.target.value = '';
                                    }
                                  }}
                                />
                              </div>
                         </div>
                     )}
                     <p className="text-xs text-muted-foreground">* Campo obrigatório.</p>
                 </div>
                 <DialogFooter>
                     <DialogClose asChild>
                         <Button type="button" variant="outline">Cancelar</Button>
                     </DialogClose>
                     {configManagementDialogState.action === 'save' && (
                         <Button type="button" onClick={saveCurrentConfig} disabled={!configManagementDialogState.configName.trim()}>Salvar Modelo</Button>
                     )}
                     {configManagementDialogState.action === 'load' && (
                         <Button type="button" onClick={loadSelectedConfig} disabled={!configManagementDialogState.selectedConfigToLoad}>Carregar Modelo</Button>
                     )}
                 </DialogFooter>
             </DialogContent>
         </Dialog>


    </div>
  );
}

// Helper: Array of value field IDs for quick check
const valorFieldIds = ['valor_parcela', 'valor_financiado', 'valor_total', 'valor_previsto', 'valor_realizado', 'margem_bruta', 'margem_reservada', 'margem_liquida'];
