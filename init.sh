#!/bin/bash -u

INIT_SH=$(basename $0)
SCRIPT_DIR=$(cd $(dirname $0) && pwd)
BIN_DIR=${SCRIPT_DIR}/bin
SRC_DIR=${SCRIPT_DIR}/src
HEADER=${INIT_SH}": "

# 現状、Cygwin のみ対応
RESULT=0
OUTPUT=$(type cygpath 2>&1 > /dev/null) || RESULT=$? 
if [ ! "$RESULT" = "0" ]; then
	echo ${INIT_SH}": command 'cygpath' not found."
	exit 1
fi

# 引数がない場合は終了
if [ $# -lt 1 ]; then
	echo ${INIT_SH}": Missing argument (a .xlsm file or new project name)" >&2
	exit 1
fi

WSF_WINDOWS_PATH=$(cygpath -w ${SCRIPT_DIR}/vbac.wsf) # vbac.wsf のWindows形式フルパス

if [ -e $1 ]; then # .xlsm ファイルが指定された場合、bin/ にコピーしてdecombine（拡張子のチェックはしない）
	NEW_FILE=`basename $1`

	# 引数がディレクトリだったら終了
	if [ -d $1 ]; then
		echo ${INIT_SH}": "$1" is directory." >&2
		exit 1
	fi

	# bin/ に同名のファイルが既に存在したら終了
	if [ -e ${BIN_DIR}/${NEW_FILE} ]; then
		echo ${INIT_SH}": bin/"${NEW_FILE}" is already exists." >&2
		exit 1
	fi

	# src/ に同名のディレクトリが存在したら警告
	if [ -e ${SRC_DIR}/${NEW_FILE} ]; then
		echo ${INIT_SH}": src/"${NEW_FILE}" is already exists. \
		If you run \`cscript //nologo vbac.wsf decombine\`, \
		VBA source files in src/"${NEW_FILE}" will be overwrited with files in "$1"."
		echo
	fi

	# 指定されたファイルを bin/ にコピーした後、任意で decombine を実行
	cp $1 ${BIN_DIR}
	echo "About to run \`cscript //nologo vbac.wsf decombine\`. This will overwrite existing VBA source files in 'src' directory."
	read -p "Do you want to proceed? [Y/n]: " ANSWER
	case $ANSWER in
		[Yy]* )
			cscript //nologo ${WSF_WINDOWS_PATH} decombine
			;;
		*)
			echo "Running \`vbac.wsf\` was skipped."
			echo "You can run \`cscript //nologo vbac.wsf decombine\` to export VBA source files from '"$1".xlsm'."
			exit 0
			;;
	esac
else # ファイルが存在しない場合、プロジェクト名として扱い、xlsmファイルをcombineで作成する
	NEW_FILE=$1".xlsm"

	# プロジェクト名に '/' は使用不能
	if [ $(echo "${NEW_FILE}" | grep "/") ]; then
		echo ${INIT_SH}": The project name cannot contain slashes('/')."
		exit 1
	fi

	# .xslm を付けて、src/ に同名のディレクトリが既に存在したら終了
	if [ -e ${SRC_DIR}/${NEW_FILE} ]; then
		echo ${INIT_SH}": "src/${NEW_FILE}" is already exists." >&2
		exit 1
	fi

	# .xslm を付けて、bin/ に同名のファイルが既に存在したら終了
	if [ -e ${SRC_DIR}/${NEW_FILE} ]; then
		echo ${INIT_SH}": "bin/${NEW_FILE}" is already exists." >&2
		exit 1
	fi

	# src/ にディレクトリを作成した後、任意で decombine を実行
	mkdir ${SRC_DIR}/${NEW_FILE}
	echo "About to run \`cscript //nologo vbac.wsf combine\`. This will overwrite existing VBA source files in .xlsm files."
	read -p "Do you want to proceed? [Y/n]: " ANSWER
	case $ANSWER in
		[Yy]* )
			cscript //nologo ${WSF_WINDOWS_PATH} combine
			;;
		*)
			echo "Running vbac.wsf was skipped."
			echo "You can run \`cscript //nologo vbac.wsf combine\` to generate '${NEW_FILE}'."
			exit 0
			;;
	esac
fi