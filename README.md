# ppt_xfr
�쐬�����p���[�|�C���g�̃X���C�h���A�폜�A�\��/��\���؂�ւ��A�m�[�g�폜�����邱�Ƃ��ł��܂��B

# �����
* OS:Windows
* Microsoft PowerPoint
* .ppt/.pptx

Windows7 + PowerPoint 2010�œ���m�F�����Ă��܂��B

������`.ppt`�̏ꍇ�A��q����悤��`/o`�I�v�V�����𖾎�����K�v������܂��B

# �g����
```
> git clone https://github.com/MATSUMOTO-Takashi/ppt_xfr.git
> cd ppt_xfr
> main.bat /f:config_example.json /o:output.pptx test.pptx
```

���s���̏����F  
`main.bat [/f] [/o] ���t�@�C��`

# ���s���I�v�V����
���s���Ɉȉ��̃I�v�V�������w�肷�邱�Ƃ��ł��܂��B

## /f
�w�肵���R���t�B�O�t�@�C�����g�p���܂��B

�ȗ������ꍇ`config.json`��ǂݍ��݂܂��B

## /o
�o�̓t�@�C���̃t�@�C�������w�肵�܂��B

�ȗ������ꍇ`output.pptx`���o�͂��܂��B

**���ӁI**�F���t�@�C����2007�ȑO��`.ppt`�̏ꍇ�A�����I��`/o`���w�肵�Ȃ��ƃG���[�ƂȂ�܂��B

## /h
�w���v��\�����܂��B

# �ϊ��p�ݒ�t�@�C��
ex) config_example.json

```
{
  "slide": {
    "show": -1,
    "hidden": [],
    "del": [3, 5, 24]
  },
  "note": {
    "del": -1
  }
}
```

## slide.show
�X���C�h�V���[���̕\���ݒ�ł��B

type: -1 or array

-1��ݒ肷��ƑS�ẴX���C�h�ɑ΂��ĕ\���ݒ���s���܂��B

�w�肵���X���C�h�̂ݕ\���ݒ���s�������ꍇ��`[1, 2, 3]`�̂悤�ɔz��^���g�p���A�X���C�h�ԍ��𗅗񂵂Ă��������B

## slide.hidden
�X���C�h�V���[���̔�\���ݒ�ł��B

type: array

�w�肵���X���C�h�̔�\���ݒ�����܂��B`[1, 2, 3]`�̂悤�ɔz��^���g�p���A�X���C�h�ԍ��𗅗񂵂Ă��������B

slide.show��slide.hidden�œ����ԍ����w�肳�ꂽ�ꍇ�A**��\�����**�ƂȂ�܂��B

## slide.del
�X���C�h�̍폜�����܂��B

type: array

�w�肵���X���C�h���폜���܂��B`[1, 2, 3]`�̂悤�ɔz��^���g�p���A�X���C�h�ԍ��𗅗񂵂Ă��������B

## note.del
�m�[�g�̍폜�����܂��B

type: -1 or array

-1��ݒ肷��ƑS�ẴX���C�h�̃m�[�g���폜���܂��B

�w�肵���X���C�h�̃m�[�g�݂̂��폜�������ꍇ��`[1, 2, 3]`�̂悤�ɔz��^���g�p���A�X���C�h�ԍ��𗅗񂵂Ă��������B

## �X���C�h�ԍ��̎w��ɂ���
�X���C�h�̍폜�͈�A�̑���̍Ō�Ɏ��s����܂��B

����āA�X���C�h�̔ԍ��̓p���[�|�C���g���J�����ۂɐU���Ă���X���C�h�ԍ������̂܂܎w�肵�Ă��������B
