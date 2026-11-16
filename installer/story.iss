{ Zadoo's lightweight installer companion. No animation code runs in the app. }
const
  StoryBackground = $2E2310;
  StoryInk = $FFF8ED;
  StoryCyan = $EADB60;
  StoryMuted = $C9BAA2;
  SPI_GETCLIENTAREAANIMATION = $1042;

function SetTimer(hWnd: HWND; nIDEvent: UINT_PTR; uElapse: UINT; lpTimerFunc: NativeInt): UINT_PTR;
  external 'SetTimer@user32.dll stdcall';
function KillTimer(hWnd: HWND; uIDEvent: UINT_PTR): BOOL;
  external 'KillTimer@user32.dll stdcall';
function GetForegroundWindow(): HWND;
  external 'GetForegroundWindow@user32.dll stdcall';
function ReadAnimationPreference(Action: UINT; Param: UINT; var Value: BOOL; Flags: UINT): BOOL;
  external 'SystemParametersInfoW@user32.dll stdcall';

var
  StoryTimer: UINT_PTR;
  StoryTick: Integer;
  StoryActiveScene: Integer;
  StoryBasePose: Integer;
  StoryClips: array[0..2] of TPanel;
  StorySprites: array[0..2] of TBitmapImage;
  StoryClipTops: array[0..2] of Integer;
  StoryMotion: TNewCheckBox;
  StoryStep: TNewStaticText;
  StoryInstallTitle: TNewStaticText;
  StoryInstallDetail: TNewStaticText;
  StoryLastProgress: Integer;
  StoryLastMilestone: Integer;

function AddStoryText(Parent: TWinControl; Text: String; X, Y, W, H, Size: Integer; Ink: TColor): TNewStaticText;
begin
  Result := TNewStaticText.Create(WizardForm);
  Result.Parent := Parent;
  Result.SetBounds(X, Y, W, H);
  Result.AutoSize := False;
  Result.WordWrap := True;
  Result.ShowAccelChar := False;
  Result.StyleElements := [];
  Result.Color := StoryBackground;
  Result.Font.Name := 'Segoe UI';
  Result.Font.Size := Size;
  Result.Font.Color := Ink;
  Result.Caption := Text;
end;

function AddStoryPanel(Parent: TWinControl; X, Y, W, H: Integer): TPanel;
begin
  Result := TPanel.Create(WizardForm);
  Result.Parent := Parent;
  Result.SetBounds(X, Y, W, H);
  Result.StyleElements := [];
  Result.Color := StoryBackground;
  Result.BevelOuter := bvNone;
  Result.BevelInner := bvNone;
end;

procedure SetStoryPose(Scene, Pose: Integer);
begin
  StorySprites[Scene].Left := -(Pose mod 2) * StoryClips[Scene].Width;
  StorySprites[Scene].Top := -(Pose div 2) * StoryClips[Scene].Height;
end;

procedure AddStorySprite(Parent: TWinControl; Scene, X, Y, Size, Pose: Integer);
begin
  StoryClips[Scene] := AddStoryPanel(Parent, X, Y, Size, Size);
  StoryClipTops[Scene] := Y;
  StorySprites[Scene] := TBitmapImage.Create(WizardForm);
  StorySprites[Scene].Parent := StoryClips[Scene];
  StorySprites[Scene].SetBounds(0, 0, Size * 2, Size * 2);
  StorySprites[Scene].BackColor := StoryBackground;
  StorySprites[Scene].Stretch := True;
  StorySprites[Scene].PngImage.LoadFromFile(ExpandConstant('{tmp}\mascot-poses.png'));
  SetStoryPose(Scene, Pose);
end;

procedure StopStoryMotion();
var
  I: Integer;
begin
  if StoryTimer <> 0 then
  begin
    KillTimer(0, StoryTimer);
    StoryTimer := 0;
  end;
  for I := 0 to 2 do
    if StoryClips[I] <> nil then StoryClips[I].Top := StoryClipTops[I];
  if StoryActiveScene >= 0 then
    if StorySprites[StoryActiveScene] <> nil then
      SetStoryPose(StoryActiveScene, StoryBasePose);
end;

procedure StoryTimerTick(Window: HWND; Message: UINT; Event: UINT_PTR; Time: DWORD);
var
  Phase, Offset, Pose: Integer;
begin
  Inc(StoryTick);
  #ifdef ZadooVisualPreview
  if StoryTick = 1 then Log('Zadoo story: timer started (80 ms interval)');
  if StoryTick = 100 then Log('Zadoo story: bounded animation completed');
  #endif
  { Eight-second bursts only. Hidden/inactive windows do not redraw. }
  if (StoryTick >= 100) or (not StoryMotion.Checked) then
  begin
    StopStoryMotion();
    Exit;
  end;
  if (StoryActiveScene < 0) or (GetForegroundWindow() <> WizardForm.Handle) then Exit;
  Phase := StoryTick mod 24;
  if Phase > 12 then Phase := 24 - Phase;
  Offset := (Phase div 3) - 2;
  StoryClips[StoryActiveScene].Top := StoryClipTops[StoryActiveScene] + ScaleY(Offset);
  Pose := StoryBasePose;
  if (StoryActiveScene = 0) and ((StoryTick mod 36) >= 33) then Pose := 1;
  SetStoryPose(StoryActiveScene, Pose);
  #ifdef ZadooVisualPreview
  if (StoryTick = 8) or (StoryTick = 24) then
    Log('Zadoo story: rendered scene=' + IntToStr(StoryActiveScene) + ' pose=' + IntToStr(Pose) + ' offset=' + IntToStr(Offset));
  #endif
end;

procedure StartStoryMotion();
begin
  StopStoryMotion();
  StoryTick := 0;
  if (StoryActiveScene >= 0) and StoryMotion.Checked and (not WizardSilent()) then
    StoryTimer := SetTimer(0, 0, 80, CreateCallback(@StoryTimerTick));
end;

procedure StoryMotionChanged(Sender: TObject);
begin
  StartStoryMotion();
end;

procedure AddStorySide(Parent: TWinControl; Reference: TBitmapImage; Scene, Pose: Integer; Message: String);
var
  Panel: TPanel;
  Title: TNewStaticText;
  SpriteY: Integer;
begin
  Reference.Visible := False;
  Panel := AddStoryPanel(Parent, Reference.Left, Reference.Top, Reference.Width, Reference.Height);
  Title := AddStoryText(Panel, 'zadoo', ScaleX(18), ScaleY(18), Panel.Width - ScaleX(36), ScaleY(32), 22, StoryInk);
  Title.Font.Style := [fsBold];
  SpriteY := ScaleY(65);
  AddStorySprite(Panel, Scene, 0, SpriteY, Panel.Width, Pose);
  Title := AddStoryText(Panel, Message, ScaleX(14), SpriteY + Panel.Width + ScaleY(8), Panel.Width - ScaleX(28), ScaleY(42), 10, StoryInk);
  Title.Alignment := taCenter;
  Title := AddStoryText(Panel, 'YOUR SETUP BUDDY', ScaleX(8), Panel.Height - ScaleY(23), Panel.Width - ScaleX(16), ScaleY(16), 8, StoryCyan);
  Title.Alignment := taCenter;
end;

procedure AddWelcomeSteps();
var
  I, X, Y, W: Integer;
  Card: TPanel;
  StepLabel: TNewStaticText;
  Titles: array[0..2] of String;
begin
  Titles[0] := 'Choose';
  Titles[1] := 'Install';
  Titles[2] := 'Connect';
  X := WizardForm.WelcomeLabel2.Left;
  Y := WizardForm.WelcomePage.ClientHeight - ScaleY(84);
  W := (WizardForm.WelcomeLabel2.Width - ScaleX(16)) div 3;
  WizardForm.WelcomeLabel2.Height := Y - WizardForm.WelcomeLabel2.Top - ScaleY(12);
  for I := 0 to 2 do
  begin
    Card := AddStoryPanel(WizardForm.WelcomePage, X + I * (W + ScaleX(8)), Y, W, ScaleY(66));
    StepLabel := AddStoryText(Card, '0' + IntToStr(I + 1), ScaleX(12), ScaleY(8), W - ScaleX(24), ScaleY(18), 9, StoryCyan);
    StepLabel.Font.Style := [fsBold];
    StepLabel := AddStoryText(Card, Titles[I], ScaleX(12), ScaleY(31), W - ScaleX(24), ScaleY(22), 11, StoryInk);
    StepLabel.Font.Style := [fsBold];
  end;
end;

procedure InitializeWizard();
var
  AllowMotion: BOOL;
  Panel: TPanel;
  Top, Height, Size, TextY: Integer;
begin
  StoryActiveScene := -1;
  StoryLastProgress := -1;
  StoryLastMilestone := -20;
  if WizardSilent() then Exit;
  ExtractTemporaryFile('mascot-poses.png');
  AddStorySide(WizardForm.WelcomePage, WizardForm.WizardBitmapImage, 0, 0,
    'Hey! Let''s get your PC ready.');
  AddStorySide(WizardForm.FinishedPage, WizardForm.WizardBitmapImage2, 2, 3,
    'All set. Your next stop: connect!');
  WizardForm.WelcomeLabel1.StyleElements := [seClient, seBorder];
  WizardForm.WelcomeLabel1.Font.Size := 18;
  WizardForm.WelcomeLabel1.Font.Color := StoryInk;
  WizardForm.WelcomeLabel1.Height := ScaleY(52);
  AddWelcomeSteps();

  Top := WizardForm.FilenameLabel.Top + WizardForm.FilenameLabel.Height;
  if Top < WizardForm.ProgressGauge.Top + WizardForm.ProgressGauge.Height then
    Top := WizardForm.ProgressGauge.Top + WizardForm.ProgressGauge.Height;
  Top := Top + ScaleY(18);
  Height := WizardForm.InstallingPage.ClientHeight - Top - ScaleY(4);
  Size := ScaleX(180);
  if Size > Height - ScaleY(24) then Size := Height - ScaleY(24);
  Panel := AddStoryPanel(WizardForm.InstallingPage, 0, Top, WizardForm.InstallingPage.ClientWidth, Height);
  AddStorySprite(Panel, 1, ScaleX(8), (Height - Size) div 2, Size, 2);
  TextY := (Height - ScaleY(116)) div 2;
  StoryInstallTitle := AddStoryText(Panel, 'Building your workspace.', Size + ScaleX(24), TextY,
    Panel.Width - Size - ScaleX(42), ScaleY(48), 16, StoryInk);
  StoryInstallTitle.Font.Style := [fsBold];
  StoryInstallDetail := AddStoryText(Panel, 'Your setup buddy is putting everything in place.', Size + ScaleX(24), TextY + ScaleY(55),
    Panel.Width - Size - ScaleX(42), ScaleY(48), 10, StoryMuted);

  StoryMotion := TNewCheckBox.Create(WizardForm);
  StoryMotion.Parent := WizardForm;
  StoryMotion.Caption := 'Animations';
  StoryMotion.SetBounds(ScaleX(16), WizardForm.NextButton.Top, ScaleX(94), WizardForm.NextButton.Height);
  StoryMotion.Anchors := [akLeft, akBottom];
  StoryMotion.Hint := 'Optional character motion. Starts with your Windows animation preference.';
  AllowMotion := False;
  if ReadAnimationPreference(SPI_GETCLIENTAREAANIMATION, 0, AllowMotion, 0) then
    StoryMotion.Checked := AllowMotion;
  StoryMotion.OnClick := @StoryMotionChanged;
  StoryStep := AddStoryText(WizardForm, '01 / HELLO', ScaleX(120), WizardForm.NextButton.Top + ScaleY(5),
    ScaleX(150), ScaleY(20), 8, StoryCyan);
  StoryStep.Color := WizardForm.Color;
  StoryStep.Anchors := [akLeft, akBottom];
  #ifdef ZadooVisualPreview
  WizardForm.Caption := 'Zadoo · Cartoon preview (no installation)';
  #endif
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if WizardSilent() then Exit;
  #ifdef ZadooVisualPreview
  { Preview only: display these pages without executing installation operations. }
  if (CurPageID = wpWelcome) and (ExpandConstant('{param:StoryPreview|welcome}') = 'installing') then
  begin
    WizardForm.OuterNotebook.ActivePage := WizardForm.InnerPage;
    WizardForm.InnerNotebook.ActivePage := WizardForm.InstallingPage;
    WizardForm.PageNameLabel.Caption := 'Installing Zadoo · visual preview';
    WizardForm.PageDescriptionLabel.Caption := 'No files or settings are being installed.';
    WizardForm.StatusLabel.Caption := 'Preview of the real progress layout';
    WizardForm.FilenameLabel.Caption := 'No installation operations are available in this preview.';
    WizardForm.NextButton.Enabled := False;
    CurPageID := wpInstalling;
  end;
  if (CurPageID = wpWelcome) and (ExpandConstant('{param:StoryPreview|welcome}') = 'finished') then
  begin
    WizardForm.OuterNotebook.ActivePage := WizardForm.FinishedPage;
    WizardForm.FinishedLabel.Caption := SetupMessage(msgFinishedLabel);
    WizardForm.NextButton.Enabled := False;
    CurPageID := wpFinished;
  end;
  #endif
  StopStoryMotion();
  StoryActiveScene := -1;
  StoryBasePose := 0;
  if CurPageID = wpWelcome then
  begin
    StoryActiveScene := 0;
    StoryStep.Caption := '01 / HELLO';
  end
  else if CurPageID = wpInstalling then
  begin
    StoryActiveScene := 1;
    StoryBasePose := 2;
    StoryStep.Caption := '03 / MAKING IT YOURS';
  end
  else if CurPageID = wpFinished then
  begin
    StoryActiveScene := 2;
    StoryBasePose := 3;
    StoryStep.Caption := '04 / LET''S CONNECT';
    { Long, plain-language completion instructions must never be clipped. }
    WizardForm.AdjustLabelHeight(WizardForm.FinishedLabel);
    WizardForm.RunList.Top := WizardForm.FinishedLabel.Top + WizardForm.FinishedLabel.Height + ScaleY(12);
    WizardForm.RunList.Height := WizardForm.FinishedPage.ClientHeight - WizardForm.RunList.Top - ScaleY(12);
  end
  else StoryStep.Caption := '02 / YOUR CHOICES';
  StartStoryMotion();
end;

procedure CurInstallProgressChanged(CurProgress, MaxProgress: Integer);
var
  Percent: Integer;
begin
  if WizardSilent() or (MaxProgress <= 0) then Exit;
  Percent := Round((CurProgress * 100.0) / MaxProgress);
  if Percent = StoryLastProgress then Exit;
  StoryLastProgress := Percent;
  StoryInstallTitle.Caption := 'Settling in... ' + IntToStr(Percent) + '%';
  if Percent < 35 then StoryInstallDetail.Caption := 'Unpacking the app files.'
  else if Percent < 85 then StoryInstallDetail.Caption := 'Putting your workspace in place.'
  else StoryInstallDetail.Caption := 'Finishing the file copy. Nearly there!';
  if Percent >= StoryLastMilestone + 20 then
  begin
    StoryLastMilestone := Percent;
    StartStoryMotion();
  end;
end;

procedure DeinitializeSetup();
begin
  StopStoryMotion();
end;
