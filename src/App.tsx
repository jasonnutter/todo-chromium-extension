import React, { useState, useRef } from "react";
import "./App.css";
import { AccountInfo } from "@azure/msal-browser";
import {
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  CompoundButton,
  Persona,
  PersonaSize,
  Pivot,
  PivotItem,
  TextField,
  ComboBox,
  IComboBoxOption,
  IComboBox,
  Label,
} from "@fluentui/react";

import { initializeIcons } from "@fluentui/react/lib/Icons";

import { logout, login, getActiveAccount } from "./auth";
import { useSyncedValue, useCurrentTab, useSignedInUser } from "./chrome";
import { useGraphProfile, addTask, getTaskFolders } from "./graph";

initializeIcons();

const DEFAULT_TASKS_FOLDER_NAME = "Links";
const TASK_FOLDER_SYNC_KEY = "taskFolder";
const NEW_FOLDER_SUFFIX = " (New)";

const Account: React.FC = () => {
  const [graphProfile] = useGraphProfile({
    userPrincipalName: "",
    displayName: "",
  });

  if (!graphProfile.userPrincipalName) {
    return <Spinner size={SpinnerSize.large} />;
  }

  return (
    <Persona
      size={PersonaSize.size48}
      text={graphProfile.displayName}
      secondaryText={graphProfile.userPrincipalName}
      imageInitials={graphProfile.displayName[0]}
    />
  );
};

const App: React.FC = () => {
  const [account, setAccount] = useState<AccountInfo | null>(
    getActiveAccount()
  );
  const [success, setSuccess] = useState<boolean>(false);
  const [error, setError] = useState<string>("");
  const [inProgress, setInProgress] = useState<boolean>(false);
  const [latestTask, setLatestTask] = useState<string | undefined>("");
  const [selectedTaskFolderIndex, setSelectedTaskFolderIndex] =
    useState<number>(0);

  // Selected folder (in progress)
  const [selectedTaskFolder, setSelectedTaskFolder] =
    useState<IComboBoxOption | null>(null);

  // Folder name (saved)
  const [, setSavedTaskFolder] = useState<string>(DEFAULT_TASKS_FOLDER_NAME);

  // Folders fetched from API
  const [taskFolders, setTaskFolders] = useState<IComboBoxOption[]>([
    { key: DEFAULT_TASKS_FOLDER_NAME, text: DEFAULT_TASKS_FOLDER_NAME },
  ]);

  // Load synced folder name
  const [, setSyncedFolderName] = useSyncedValue<string>(
    TASK_FOLDER_SYNC_KEY,
    DEFAULT_TASKS_FOLDER_NAME,
    (folderName) => {
      const taskFolder = {
        key: folderName,
        text: folderName,
      };

      setTaskFolders([taskFolder]);
      setSavedTaskFolder(folderName);
      setSelectedTaskFolder(taskFolder);
    }
  );

  const comboBoxRef = useRef<IComboBox | null>(null);

  const [currentTab, setCurrentTab] = useCurrentTab({
    title: "",
    url: "",
  });

  const [signedInUser] = useSignedInUser({
    email: "",
    id: "",
  });

  return (
    <div className="wrapper">
      {account ? (
        <Pivot>
          <PivotItem headerText="Save Link">
            <Stack tokens={{ childrenGap: 15 }}>
              <form
                onSubmit={async (e) => {
                  e.preventDefault();

                  setSuccess(false);
                  setInProgress(true);
                  setLatestTask("");

                  if (selectedTaskFolder && selectedTaskFolder.text) {
                    const newFolderName = selectedTaskFolder.text
                      .split(NEW_FOLDER_SUFFIX)[0]
                      .trim();

                    setSavedTaskFolder(newFolderName);
                    setSyncedFolderName(newFolderName);

                    try {
                      const { id } = await addTask(
                        currentTab.title || "",
                        currentTab.url || "",
                        newFolderName
                      );
                      setLatestTask(id);
                      setSuccess(true);
                      setError("");
                    } catch (e) {
                      setError(e.message);
                      setSuccess(false);
                    }
                  }

                  setInProgress(false);
                }}
              >
                <ul>
                  <li>
                    <TextField
                      label="Title"
                      value={currentTab.title}
                      onChange={(e) => {
                        const target = e.target as HTMLTextAreaElement;

                        setCurrentTab({
                          ...currentTab,
                          title: target.value,
                        });
                      }}
                    />
                  </li>
                  <li>
                    <TextField
                      label="URL"
                      multiline={true}
                      value={currentTab.url}
                      onChange={(e) => {
                        const target = e.target as HTMLInputElement;

                        setCurrentTab({
                          ...currentTab,
                          url: target.value,
                        });
                      }}
                    />
                  </li>
                  <li>
                    <ComboBox
                      allowFreeform={true}
                      selectedKey={
                        selectedTaskFolder
                          ? selectedTaskFolder.key
                          : taskFolders[0].key
                      }
                      label="Task Folder"
                      onPendingValueChanged={async (option, index, value) => {
                        if (value) {
                          const folders = await getTaskFolders(value);
                          const folderExists = folders.value.find(
                            (folder) => folder.name === value
                          );
                          const newFolders: IComboBoxOption[] = folderExists
                            ? []
                            : [
                                {
                                  key: value,
                                  text: `${value}${NEW_FOLDER_SUFFIX}`,
                                },
                              ];

                          if (comboBoxRef.current) {
                            comboBoxRef.current.focus(true);
                          }

                          if (folders.value.length) {
                            setTaskFolders(
                              newFolders.concat(
                                folders.value.map((folder) => ({
                                  key: folder.id,
                                  text: folder.name,
                                }))
                              )
                            );
                          } else {
                            setTaskFolders(newFolders);
                          }
                        } else {
                          setSelectedTaskFolder(null);
                        }
                      }}
                      onItemClick={(e, option, index) => {
                        if (option && option.text) {
                          setSelectedTaskFolder(option);
                          setSelectedTaskFolderIndex(index || 0);
                        }
                      }}
                      onChange={(e, option, index) => {
                        const target = e.target as HTMLInputElement;

                        if (option && option.text) {
                          setSelectedTaskFolder(option);
                          setSelectedTaskFolderIndex(index || 0);
                        } else {
                          const folderIndex = taskFolders.findIndex(
                            (folder) =>
                              folder.text.split(NEW_FOLDER_SUFFIX)[0] ===
                              target.value
                          );
                          setSelectedTaskFolder(taskFolders[folderIndex]);
                          setSelectedTaskFolderIndex(folderIndex);
                        }
                      }}
                      onBlur={(e) => {
                        if (taskFolders.length) {
                          setSelectedTaskFolder(
                            taskFolders[selectedTaskFolderIndex]
                          );
                        }
                      }}
                      onScrollToItem={(index) => {
                        setSelectedTaskFolderIndex(index > -1 ? index : 0);
                      }}
                      componentRef={comboBoxRef}
                      options={taskFolders}
                    />
                  </li>
                  <li>
                    <PrimaryButton
                      disabled={
                        inProgress || !selectedTaskFolder || !currentTab.title
                      }
                      type="submit"
                    >
                      Save Link
                    </PrimaryButton>
                  </li>
                </ul>
              </form>

              {(inProgress || success || error) && (
                <Stack tokens={{ childrenGap: 15 }}>
                  {inProgress && <Spinner size={SpinnerSize.medium} />}
                  {success && (
                    <MessageBar messageBarType={MessageBarType.success}>
                      Link saved successfully.
                    </MessageBar>
                  )}
                  {error && (
                    <MessageBar messageBarType={MessageBarType.error}>
                      {error}
                    </MessageBar>
                  )}
                  {latestTask && (
                    <DefaultButton
                      href={`https://to-do.microsoft.com/tasks/id/${latestTask}/details`}
                      target="_blank"
                    >
                      View Task
                    </DefaultButton>
                  )}
                </Stack>
              )}
            </Stack>
          </PivotItem>
          <PivotItem headerText="Account">
            <Stack tokens={{ childrenGap: 15 }}>
              <Label>Account</Label>

              {account && <Account />}

              <DefaultButton
                onClick={async () => {
                  setAccount(null);
                  try {
                    await logout();
                    setError("");
                  } catch (e) {
                    setError(e.message);
                  }
                }}
              >
                Logout
              </DefaultButton>
            </Stack>
          </PivotItem>
        </Pivot>
      ) : (
        <Pivot>
          <PivotItem headerText="Account" style={{ paddingTop: "15px" }}>
            <Stack horizontal={true} tokens={{ childrenGap: 15 }}>
              {signedInUser.email && (
                <CompoundButton
                  primary={true}
                  onClick={async () => {
                    try {
                      await login(signedInUser.email);
                      setAccount(getActiveAccount());
                      setError("");
                    } catch (e) {
                      setError(e.message);
                    }
                  }}
                  secondaryText={`(w/ ${signedInUser.email})`}
                  style={{
                    width: "200px",
                  }}
                >
                  Login
                </CompoundButton>
              )}
              <CompoundButton
                onClick={async () => {
                  try {
                    await login();
                    setAccount(getActiveAccount());
                    setError("");
                  } catch (e) {
                    setError(e.message);
                  }
                }}
                secondaryText="(w/ your Microsoft account)"
                style={{
                  width: "200px",
                }}
              >
                Login
              </CompoundButton>
            </Stack>
            {error && (
              <div style={{ marginTop: "15px" }}>
                <MessageBar messageBarType={MessageBarType.error}>
                  {error}
                </MessageBar>
              </div>
            )}
          </PivotItem>
        </Pivot>
      )}
    </div>
  );
};

export default App;
