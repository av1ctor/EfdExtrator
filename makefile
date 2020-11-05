FBC := fbc

rootdir := $(dir $(MAKEFILE_LIST))
srcdir := $(rootdir)
objdir := $(rootdir)obj

APP_EXE := EfdExtrator.exe
APP_BI  := $(wildcard $(srcdir)/*.bi)
APP_BAS := $(sort $(wildcard $(srcdir)/*.bas))
APP_OBJ := $(patsubst $(srcdir)/%.bas,$(objdir)/%.o,$(APP_BAS))

LIBS_BI  := $(wildcard $(srcdir)/libs/*.bi)
LIBS_BAS := $(sort $(wildcard $(srcdir)/libs/*.bas))
LIBS_OBJ := $(patsubst $(srcdir)/libs/%.bas,$(objdir)/%.o,$(LIBS_BAS))

APP_FLAGS := gui.rc -x $(APP_EXE)
OBJ_FLAGS := -d WITH_PARSER -O 3

.PHONY: app
app: $(APP_EXE)

$(APP_EXE): $(LIBS_OBJ) $(APP_OBJ)
	$(FBC) $(APP_FLAGS) $^

$(APP_OBJ): $(objdir)/%.o: $(srcdir)/%.bas $(APP_BI) $(LIBS_BI) | $(objdir)
	$(FBC) $(OBJ_FLAGS) -m EfdMain -c $< -o $@

$(LIBS_OBJ): $(objdir)/%.o: $(srcdir)/libs/%.bas $(LIBS_BI) | $(objdir)
	$(FBC) $(OBJ_FLAGS) -c $< -o $@

