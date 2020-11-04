FBC := fbc

rootdir := $(dir $(MAKEFILE_LIST))
srcdir := $(rootdir)
objdir := $(rootdir)obj

APP_EXE := EfdExtrator.exe
APP_BI  :=        $(wildcard $(srcdir)/*.bi)
APP_BAS := $(sort $(wildcard $(srcdir)/*.bas))
APP_BAS := $(patsubst $(srcdir)/%.bas,$(objdir)/%.o,$(APP_BAS))

APP_FLAGS := gui.rc -x $(APP_EXE)
OBJ_FLAGS := -m EfdMain -d WITH_PARSER -O 3

.PHONY: app
app: $(APP_EXE)

$(APP_EXE): $(APP_BAS)
	$(FBC) $(APP_FLAGS) $^

$(APP_BAS): $(objdir)/%.o: $(srcdir)/%.bas $(APP_BI) | $(objdir)
	$(FBC) $(OBJ_FLAGS) -c $< -o $@

