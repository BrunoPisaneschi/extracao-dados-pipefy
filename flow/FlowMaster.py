# -*- coding:utf-8 -*-
# /usr/bin/python3


import re
from unidecode import unidecode
from utils.Pipefy import Pipefy
from utils.ExcelUtils import excel

class pipefy():
    def __init__(self, token, caminho_db, nome_db, caminho_excel, nome_excel):
        self._token = token
        self._pipe_id = 1102385
        self._dict_datas = {}
        self._excel_utils = excel(caminho_excel, nome_excel)
        self._db_utils = excel(caminho_db, nome_db)

    def extract_datas(self):
        _pipefy = Pipefy(self._token)
        _pipes = _pipefy.pipes([self._pipe_id])[0]
        _count = 1
        for _cards in _pipes['phases']:
            if 'resposta recebida' in _cards['name'].lower():
                for _edge in _cards['cards']['edges']:
                    self._dict_datas[_count] = {'id_card': '',
                                               'titulo': '',
                                               'vaga': '',
                                               'nivel': '',
                                               'motivo_recusa': '',
                                               'soube_vaga': ''}
                    print("Id do card: {}".format(_edge['node']['id']))
                    print("Titulo do card: {}".format(_edge['node']['title']))
                    self._dict_datas[_count].update({'id_card': _edge['node']['id'], 'titulo': _edge['node']['title']})
                    _infos_card = _pipefy.card(_edge['node']['id'])
                    for _fields in _infos_card['fields']:
                        if ('selecionar vaga' in _fields['name'].lower()):
                            print("Vaga: {}".format(re.sub(r'[^a-zA-Z\s0-9]', '', _fields['value'].strip())))
                            self._dict_datas[_count].update({'vaga': re.sub(r'[^a-zA-Z\s0-9]', '', _fields['value'].strip())})

                        elif ('nível' in _fields['name'].lower()):
                            # print("Nível: {}".format(_fields['value'].strip()))
                            self._dict_datas[_count].update({'nivel': _fields['value'].strip()})

                        elif ('motivo da recusa' in _fields['name'].lower()):
                            print("Motivos da recusa: {}".format(_fields['value'].strip()))
                            self._dict_datas[_count].update({'motivo_recusa': _fields['value'].strip()})

                        elif ('como soube da vaga' in _fields['name'].lower()):
                            print("Como soube da vaga: {}".format(_fields['value'].strip()))
                            self._dict_datas[_count].update({'soube_vaga': _fields['value'].strip()})
                    print("-" * 20)
                    _exist_card = self._consult_db(self._dict_datas[_count])
                    if(_exist_card):
                        del self._dict_datas[_count]
                        continue
                    else:
                        self._db_utils.update_excel(self._dict_datas[_count])
                    motivos_recusa = self._dict_datas[_count]['motivo_recusa'].split(",")
                    if(len(motivos_recusa) > 1):
                        self._dict_datas[_count].update({'motivo_recusa': re.sub(r'[^a-zA-Z\s0-9]', '', unidecode(motivos_recusa[0]).strip())})
                        for motivo_recusa in motivos_recusa[1:]:
                            _copy_dict = self._dict_datas[_count].copy()
                            _count += 1
                            self._dict_datas[_count] = _copy_dict
                            self._dict_datas[_count].update({'motivo_recusa': re.sub(r'[^a-zA-Z\s0-9]', '', unidecode(motivo_recusa).strip())})
                    _count += 1
        self._excel_utils.update_excel(self._dict_datas)

    def _consult_db(self, _data):
        _dict_db = self._db_utils.read_excel(db=True)
        try:
            for _index_db, _row_db in _dict_db.items():
                if(int(_row_db['id_card'])==int(_data['id_card'])):
                    return True
            return False
        except Exception as erro:
            print(erro)